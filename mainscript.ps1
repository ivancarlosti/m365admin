# Microsoft 365 Admin Operations
# Menu-driven PowerShell tool for common Microsoft 365 admin tasks via Microsoft Graph.

[console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Dot-source the lib functions.
. "$PSScriptRoot\lib\Common.ps1"
. "$PSScriptRoot\lib\Mailbox-CopyToShared.ps1"
. "$PSScriptRoot\lib\OneDrive-CopyToSharePoint.ps1"
. "$PSScriptRoot\lib\Calendar-Transfer.ps1"

$RequiredModules = @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Users',
    'Microsoft.Graph.Groups',
    'Microsoft.Graph.Mail',
    'Microsoft.Graph.Files',
    'Microsoft.Graph.Sites',
    'Microsoft.Graph.Calendar',
    'Microsoft.Graph.Identity.DirectoryManagement'
)

$Scopes = @(
    'User.Read.All',
    'Group.ReadWrite.All',
    'Directory.Read.All',
    'Mail.ReadWrite',
    'Mail.ReadWrite.Shared',
    'MailboxSettings.Read',
    'Files.ReadWrite.All',
    'Sites.ReadWrite.All',
    'Calendars.ReadWrite',
    'Calendars.ReadWrite.Shared'
) -join ' '

Test-RequiredModules -Modules $RequiredModules
foreach ($m in $RequiredModules) { Import-Module $m -ErrorAction Stop }

function Connect-ToTenant {
    $tenantId = Get-SelectedTenantId
    Write-Host ""
    Write-Host "Connecting to Microsoft Graph for tenant: $tenantId" -ForegroundColor Cyan
    Connect-MgGraph -TenantId $tenantId -Scopes $Scopes -NoWelcome
    $ctx = Get-MgContext
    Write-Host "Connected as: $($ctx.Account)" -ForegroundColor Green
    return $ctx
}

function Show-MainMenu {
    param($Context)
    Clear-Host
    Write-Host ""
    Write-Host "===============================================" -ForegroundColor DarkCyan
    Write-Host "    Microsoft 365 Admin Operations" -ForegroundColor Cyan
    Write-Host "===============================================" -ForegroundColor DarkCyan
    Write-Host "  Tenant: $($Context.TenantId)" -ForegroundColor Gray
    Write-Host "  Admin:  $($Context.Account)" -ForegroundColor Gray
    Write-Host "-----------------------------------------------" -ForegroundColor DarkCyan
    Write-Host "  1. Copy mailbox messages to a shared mailbox"
    Write-Host "  2. Copy OneDrive content to a new SharePoint site"
    Write-Host "  3. Transfer calendars to another account"
    Write-Host "  4. Switch tenant"
    Write-Host "  5. Exit"
    Write-Host "==============================================="  -ForegroundColor DarkCyan
    return (Read-Host "Select an option [1-5]")
}

function Invoke-MailboxCopyMenu {
    Write-Step "Copy mailbox messages to a shared mailbox"
    $src = Read-Upn -Prompt "Source user UPN"
    if (-not $src) { return }
    $tgt = Read-Upn -Prompt "Target shared mailbox UPN"
    if (-not $tgt) { return }
    Copy-MailboxToSharedMailbox -SourceUpn $src.UserPrincipalName -TargetSharedMailboxUpn $tgt.UserPrincipalName
}

function Invoke-OneDriveCopyMenu {
    Write-Step "Copy OneDrive content to a new SharePoint site"
    $src = Read-Upn -Prompt "Source user UPN (whose OneDrive to copy)"
    if (-not $src) { return }
    $owner = Read-Upn -Prompt "Owner UPN for the new SharePoint site"
    if (-not $owner) { return }
    Copy-OneDriveToSharePointSite -SourceUpn $src.UserPrincipalName -TargetSiteOwnerUpn $owner.UserPrincipalName
}

function Invoke-CalendarTransferMenu {
    Write-Step "Transfer calendars to another account"
    $src = Read-Upn -Prompt "Source user UPN"
    if (-not $src) { return }
    $tgt = Read-Upn -Prompt "Target user UPN"
    if (-not $tgt) { return }
    $includeSecondary = Confirm-Action -Prompt "Include secondary calendars?" -DefaultYes
    $reassign = Confirm-Action -Prompt "Attempt to reassign future-event ownership? (deletes source events, emails attendees)"
    Copy-CalendarEvents `
        -SourceUpn $src.UserPrincipalName `
        -TargetUpn $tgt.UserPrincipalName `
        -IncludeSecondaryCalendars:$includeSecondary `
        -AttemptOwnershipReassign:$reassign
}

# Connect once at startup.
$context = Connect-ToTenant

# Main loop.
while ($true) {
    $choice = Show-MainMenu -Context $context
    switch ($choice) {
        '1' { Invoke-MailboxCopyMenu;    Read-Host "`nPress Enter to return to the main menu" | Out-Null }
        '2' { Invoke-OneDriveCopyMenu;   Read-Host "`nPress Enter to return to the main menu" | Out-Null }
        '3' { Invoke-CalendarTransferMenu; Read-Host "`nPress Enter to return to the main menu" | Out-Null }
        '4' {
            try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
            $context = Connect-ToTenant
        }
        '5' {
            try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
            Write-Host "Disconnected. Goodbye." -ForegroundColor Cyan
            break
        }
        default { Write-Host "Invalid selection." -ForegroundColor Yellow; Start-Sleep -Seconds 1 }
    }
}
