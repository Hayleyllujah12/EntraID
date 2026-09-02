# Prerequisites — Teams meeting/attendance scripts

Applies to `Bulk-ExtractTeamsMeetingLogs-Phased.ps1` (audit log) and
`Bulk-ExtractTeamsAttendance-Phased.ps1` (attendance reports).

> Note: PSGallery is registered by default on Windows. Each script also prompts to
> install what it needs (Ensure-Module), so this list is for pre-installing / verifying.

## Base requirement

- **PowerShell 7+** (Windows PowerShell 5.1 is unreliable for the Graph auth stack).
  Check: `$PSVersionTable.PSVersion`

## 1) Attendance extractor — `Bulk-ExtractTeamsAttendance-Phased.ps1`

```powershell
Install-Module Microsoft.Graph.Authentication      -Scope CurrentUser -Force
Install-Module Microsoft.Graph.Calendar            -Scope CurrentUser -Force
Install-Module Microsoft.Graph.CloudCommunications -Scope CurrentUser -Force
```

Does **not** need `Microsoft.Graph.Users` — UPN→object-id resolution uses
`Invoke-MgGraphRequest` (part of `Microsoft.Graph.Authentication`).

## 2) One-time admin setup for app-only (only on the machine that creates the policies)

```powershell
Install-Module MicrosoftTeams           -Scope CurrentUser -Force   # New-/Grant-CsApplicationAccessPolicy (meetings/attendance)
Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force   # New-ApplicationAccessPolicy (scope Calendars.Read to a group)
```

## 3) Audit-log extractor — `Bulk-ExtractTeamsMeetingLogs-Phased.ps1`

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
```

## Verify everything

```powershell
$PSVersionTable.PSVersion        # 7.x expected
Get-PSRepository                 # PSGallery should be listed; if missing: Register-PSRepository -Default
Get-Module -ListAvailable `
  Microsoft.Graph.Authentication, Microsoft.Graph.Calendar,
  Microsoft.Graph.CloudCommunications, MicrosoftTeams, ExchangeOnlineManagement |
  Select-Object Name, Version | Sort-Object Name
```

## Permissions (attendance, app-only) — least privilege

Microsoft Graph **application** permissions (all read-only), admin-consented:

| Permission | Purpose |
|---|---|
| `Calendars.Read` | find each organizer's Teams meetings |
| `OnlineMeetings.Read.All` | resolve JoinWebUrl → onlineMeeting id |
| `OnlineMeetingArtifact.Read.All` | read attendance reports + records |
| `User.ReadBasic.All` | *optional* — only to translate UPN → object id; skip it by putting object IDs in the CSV |

Plus scope to one security group of the organizers:
- Teams: `Grant-CsApplicationAccessPolicy -Group <group> -PolicyName ...` (meetings/attendance)
- Exchange: `New-ApplicationAccessPolicy -AppId <appid> -PolicyScopeGroupId <group> -AccessRight RestrictAccess` (calendars)

Delegated mode needs no app; scopes consented at sign-in:
`User.Read, Calendars.Read, OnlineMeetings.Read, OnlineMeetingArtifact.Read.All` (self only).

## Troubleshooting

| Symptom | Fix |
|---|---|
| `No repository with the name 'PSGallery'` | `Register-PSRepository -Default` |
| `Assembly with same name is already loaded` | Graph submodule version drift — align versions (`Install-Module Microsoft.Graph -Scope CurrentUser -Force`) or uninstall-then-reinstall the submodules. **Never** `Update-Module Microsoft.Graph`. |
| `Install-Module` blocked (execution/TLS) | `Set-ExecutionPolicy -Scope Process RemoteSigned`; `[Net.ServicePointManager]::SecurityProtocol = 'Tls12'` |
| Running as admin installs to the wrong scope | Keep `-Scope CurrentUser` so the runner account (not an elevated one) has the modules. |
