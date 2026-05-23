# Microsoft 365 Admin Operations script
Menu-driven PowerShell tool for common Microsoft 365 admin operations via Microsoft Graph. Sibling project to [gwadmin](https://github.com/ivancarlosti/gwadmin) for Google Workspace.

<!-- buttons -->
[![Stars](https://img.shields.io/github/stars/ivancarlosti/m365admin?label=⭐%20Stars&color=gold&style=flat)](https://github.com/ivancarlosti/m365admin/stargazers)
[![Watchers](https://img.shields.io/github/watchers/ivancarlosti/m365admin?label=Watchers&style=flat&color=red)](https://github.com/sponsors/ivancarlosti)
[![Forks](https://img.shields.io/github/forks/ivancarlosti/m365admin?label=Forks&style=flat&color=ff69b4)](https://github.com/sponsors/ivancarlosti)
[![Downloads](https://img.shields.io/github/downloads/ivancarlosti/m365admin/total?label=Downloads&color=success)](https://github.com/ivancarlosti/m365admin/releases)
[![GitHub commit activity](https://img.shields.io/github/commit-activity/m/ivancarlosti/m365admin?label=Activity)](https://github.com/ivancarlosti/m365admin/pulse)
[![GitHub Issues](https://img.shields.io/github/issues/ivancarlosti/m365admin?label=Issues&color=orange)](https://github.com/ivancarlosti/m365admin/issues)  
[![License](https://img.shields.io/github/license/ivancarlosti/m365admin?label=License)](LICENSE)
[![GitHub last commit](https://img.shields.io/github/last-commit/ivancarlosti/m365admin?label=Last%20Commit)](https://github.com/ivancarlosti/m365admin/commits)
[![Security](https://img.shields.io/badge/Security-View%20Here-purple)](https://github.com/ivancarlosti/m365admin/security)
[![Code of Conduct](https://img.shields.io/badge/Code%20of%20Conduct-2.1-4baaaa)](https://github.com/ivancarlosti/m365admin?tab=coc-ov-file)  
[![GitHub Sponsors](https://img.shields.io/github/sponsors/ivancarlosti?label=GitHub%20Sponsors&color=ffc0cb)][sponsor]
[![Buy Me a Coffee](https://img.shields.io/badge/Buy%20Me%20a%20Coffee-ffdd00)][buymeacoffee]
<!-- endbuttons -->

## Operations

| # | Operation | What it does |
|---|---|---|
| 1 | **Copy mailbox messages to a shared mailbox** | Copies all messages from all folders of a source mailbox into a target shared mailbox, preserving folder structure. Source is not modified. Re-runs dedupe by `internetMessageId`. |
| 2 | **Copy OneDrive content to a new SharePoint site** | Provisions a new Microsoft 365 group (which creates a SharePoint site), waits for it to come online, then copies the source user's entire OneDrive into the group's document library. Source OneDrive is not modified. |
| 3 | **Transfer calendars to another account** | Copies all events (and optionally secondary calendars) from a source user to a target user. Optional ownership reassignment for future events organized by source — see limitation below. |

## Instructions
* Save the latest release and extract files locally (download [here](https://github.com/ivancarlosti/m365auditor/releases/latest))
* Update `tenantIds.txt` with your tenants (one per line)
* Run `mainscript.ps1` in PowerShell (right-click > Run with PowerShell)
* Select a tenant, authenticate with an admin account, then choose an operation from the menu
* Logs for each operation are written to `Downloads\m365admin-logs\`
* If modules need to be installed or updated, run `ADMIN-install-modules.ps1` as Administrator

## Requirements
* Windows 10+ or Windows Server 2019+
* PowerShell 5.x or 7.x
* PowerShell modules (installed via `ADMIN-install-modules.ps1`):
  * `Microsoft.Graph.Authentication`
  * `Microsoft.Graph.Users`
  * `Microsoft.Graph.Groups`
  * `Microsoft.Graph.Mail`
  * `Microsoft.Graph.Files`
  * `Microsoft.Graph.Sites`
  * `Microsoft.Graph.Calendar`
  * `Microsoft.Graph.Identity.DirectoryManagement`

## Required Microsoft Graph scopes
The script requests the following scopes on `Connect-MgGraph`. All require admin consent:

```
User.Read.All Group.ReadWrite.All Directory.Read.All
Mail.ReadWrite Mail.ReadWrite.Shared MailboxSettings.Read
Files.ReadWrite.All Sites.ReadWrite.All
Calendars.ReadWrite Calendars.ReadWrite.Shared
```

For the mailbox copy operation, the signed-in admin must have FullAccess on the target shared mailbox.

## Known limitations
* **Mailbox copy**: Microsoft Graph does not support cross-mailbox copy actions, so messages are exported as MIME from source and re-imported on target. Headers, attachments, and `internetMessageId` are preserved; `isRead` and `categories` are re-applied via a follow-up PATCH.
* **OneDrive copy**: file version history is not preserved (only the current version is copied). OneNote (`.one`) notebooks and files over 250 GB are skipped with a warning.
* **Calendar ownership reassignment**: Microsoft Graph treats `event.organizer` as immutable — there is no Graph equivalent of Google's calendar-transfer API. The opt-in reassignment in option 3 is implemented as **delete-and-recreate** for future events where the source is the organizer, which sends cancellation emails followed by new invites to all attendees, and regenerates any Teams meeting links. This is a Microsoft platform limitation, not a tool limitation. The feature is opt-in and requires an explicit confirmation prompt.

<!-- footer -->
---

## 🧑‍💻 Consulting and technical support
* For personal support and queries, please submit a new issue to have it addressed.
* For commercial related questions, please [**contact me**][ivancarlos] for consulting costs. 

| 🩷 Project support |
| :---: |
If you found this project helpful, consider [**buying me a coffee**][buymeacoffee]
|Thanks for your support, it is much appreciated!|

[cc]: https://docs.github.com/en/communities/setting-up-your-project-for-healthy-contributions/adding-a-code-of-conduct-to-your-project
[contributing]: https://docs.github.com/en/articles/setting-guidelines-for-repository-contributors
[security]: https://docs.github.com/en/code-security/getting-started/adding-a-security-policy-to-your-repository
[support]: https://docs.github.com/en/articles/adding-support-resources-to-your-project
[it]: https://docs.github.com/en/communities/using-templates-to-encourage-useful-issues-and-pull-requests/configuring-issue-templates-for-your-repository#configuring-the-template-chooser
[prt]: https://docs.github.com/en/communities/using-templates-to-encourage-useful-issues-and-pull-requests/creating-a-pull-request-template-for-your-repository
[funding]: https://docs.github.com/en/articles/displaying-a-sponsor-button-in-your-repository
[ivancarlos]: https://ivancarlos.me
[buymeacoffee]: https://buymeacoffee.com/ivancarlos
[patreon]: https://www.patreon.com/ivancarlos
[paypal]: https://icc.gg/donate
[sponsor]: https://github.com/sponsors/ivancarlosti
