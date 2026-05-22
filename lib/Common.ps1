# Shared helpers for m365 admin operations.
# Dot-sourced by mainscript.ps1.

function Write-Step {
    param([Parameter(Mandatory)][string]$Message)
    Write-Host ""
    Write-Host ("=" * 60) -ForegroundColor DarkCyan
    Write-Host "  $Message" -ForegroundColor Cyan
    Write-Host ("=" * 60) -ForegroundColor DarkCyan
}

function Confirm-Action {
    param(
        [Parameter(Mandatory)][string]$Prompt,
        [switch]$DefaultYes
    )
    $suffix = if ($DefaultYes) { "[Y/n]" } else { "[y/N]" }
    $answer = Read-Host "$Prompt $suffix"
    if ([string]::IsNullOrWhiteSpace($answer)) {
        return [bool]$DefaultYes
    }
    return $answer -match '^(y|yes)$'
}

function Get-SelectedTenantId {
    param([string]$TenantIdsFilePath = ".\tenantIds.txt")

    if (-not (Test-Path $TenantIdsFilePath)) {
        Write-Host "Tenant IDs file not found. Please create a file named 'tenantIds.txt' with the tenant IDs, one per line." -ForegroundColor Red
        Read-Host -Prompt "Press Enter to exit"
        exit
    }

    $tenantIds = @(Get-Content -Path $TenantIdsFilePath | Where-Object { $_ -and $_.Trim() -ne "" })
    if ($tenantIds.Count -eq 0) {
        Write-Host "tenantIds.txt is empty. Add at least one tenant ID (one per line)." -ForegroundColor Red
        Read-Host -Prompt "Press Enter to exit"
        exit
    }

    Write-Host ""
    Write-Host "Select a Tenant ID to connect:"
    for ($i = 0; $i -lt $tenantIds.Count; $i++) {
        Write-Host "  $($i + 1). $($tenantIds[$i])"
    }
    do {
        $selection = Read-Host "Enter the number of the Tenant ID"
        $valid = ($selection -match '^\d+$') -and ([int]$selection -ge 1) -and ([int]$selection -le $tenantIds.Count)
        if (-not $valid) { Write-Host "Invalid selection." -ForegroundColor Yellow }
    } while (-not $valid)

    return $tenantIds[[int]$selection - 1].Trim()
}

function Test-RequiredModules {
    param([Parameter(Mandatory)][string[]]$Modules)
    $missing = @()
    foreach ($m in $Modules) {
        if (-not (Get-Module -ListAvailable -Name $m)) {
            $missing += $m
        }
    }
    if ($missing.Count -gt 0) {
        Write-Host ""
        Write-Host "The following required modules are not installed:" -ForegroundColor Red
        foreach ($m in $missing) { Write-Host "  - $m" -ForegroundColor Red }
        Write-Host ""
        Write-Host "Run ADMIN-install-modules.ps1 as Administrator to install them." -ForegroundColor Yellow
        Read-Host -Prompt "Press Enter to exit"
        exit
    }
}

function Resolve-MgUser {
    param([Parameter(Mandatory)][string]$Upn)
    try {
        $user = Get-MgUser -Filter "userPrincipalName eq '$Upn'" -Property "Id,UserPrincipalName,DisplayName,Mail,UserType" -ErrorAction Stop
        if (-not $user) {
            Write-Host "  User not found: $Upn" -ForegroundColor Yellow
            return $null
        }
        return $user
    } catch {
        Write-Host "  Failed to resolve $Upn : $($_.Exception.Message)" -ForegroundColor Yellow
        return $null
    }
}

function Read-Upn {
    param(
        [Parameter(Mandatory)][string]$Prompt,
        [int]$MaxAttempts = 3
    )
    for ($i = 0; $i -lt $MaxAttempts; $i++) {
        $upn = (Read-Host $Prompt).Trim()
        if ($upn -match '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
            $user = Resolve-MgUser -Upn $upn
            if ($user) { return $user }
        } else {
            Write-Host "  '$upn' does not look like a valid UPN." -ForegroundColor Yellow
        }
    }
    Write-Host "  Too many invalid attempts; returning to menu." -ForegroundColor Yellow
    return $null
}

function Invoke-GraphWithRetry {
    param(
        [Parameter(Mandatory)][string]$Method,
        [Parameter(Mandatory)][string]$Uri,
        [object]$Body,
        [string]$ContentType = "application/json",
        [hashtable]$Headers,
        [string]$OutputFilePath,
        [int]$MaxRetries = 5
    )

    $attempt = 0
    while ($true) {
        $attempt++
        try {
            $params = @{
                Method = $Method
                Uri    = $Uri
            }
            if ($Body) {
                if ($Body -is [string]) {
                    $params['Body'] = $Body
                } else {
                    $params['Body'] = ($Body | ConvertTo-Json -Depth 20 -Compress)
                }
            }
            if ($ContentType) { $params['ContentType'] = $ContentType }
            if ($Headers)     { $params['Headers']     = $Headers }
            if ($OutputFilePath) { $params['OutputFilePath'] = $OutputFilePath }

            return Invoke-MgGraphRequest @params -ErrorAction Stop
        } catch {
            $resp = $_.Exception.Response
            $statusCode = $null
            if ($resp -and $resp.StatusCode) {
                $statusCode = [int]$resp.StatusCode
            }
            $retryable = ($statusCode -eq 429) -or ($statusCode -ge 500 -and $statusCode -lt 600)
            if (-not $retryable -or $attempt -ge $MaxRetries) {
                throw
            }
            $retryAfter = 0
            try {
                $h = $resp.Headers['Retry-After']
                if ($h) { $retryAfter = [int]$h }
            } catch { }
            if ($retryAfter -le 0) {
                $retryAfter = [int][Math]::Pow(2, $attempt)
            }
            Write-Host "  Graph request returned $statusCode; waiting $retryAfter s (attempt $attempt/$MaxRetries)..." -ForegroundColor DarkYellow
            Start-Sleep -Seconds $retryAfter
        }
    }
}

function Get-AllPaged {
    param(
        [Parameter(Mandatory)][string]$Uri,
        [int]$MaxItems = [int]::MaxValue
    )
    $results = @()
    $next = $Uri
    while ($next -and $results.Count -lt $MaxItems) {
        $resp = Invoke-GraphWithRetry -Method GET -Uri $next
        if ($resp.value) {
            $results += $resp.value
        }
        $next = $resp.'@odata.nextLink'
    }
    return $results
}

function New-OperationLog {
    param([Parameter(Mandatory)][string]$OperationName)
    $logDir = Join-Path $env:USERPROFILE "Downloads\m365admin-logs"
    if (-not (Test-Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }
    $ts = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $logPath = Join-Path $logDir "$OperationName-$ts.log"
    "[$([DateTime]::Now.ToString('s'))] $OperationName started" | Out-File -FilePath $logPath -Encoding UTF8
    return $logPath
}

function Write-LogLine {
    param(
        [Parameter(Mandatory)][string]$LogPath,
        [Parameter(Mandatory)][string]$Message
    )
    "[$([DateTime]::Now.ToString('s'))] $Message" | Out-File -FilePath $LogPath -Append -Encoding UTF8
}

function Get-SafeFileName {
    param([Parameter(Mandatory)][string]$Name)
    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    $sb = New-Object System.Text.StringBuilder
    foreach ($c in $Name.ToCharArray()) {
        if ($invalid -contains $c) { [void]$sb.Append('_') } else { [void]$sb.Append($c) }
    }
    return $sb.ToString()
}
