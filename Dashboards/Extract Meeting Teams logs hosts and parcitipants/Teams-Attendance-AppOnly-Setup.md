# App-only (certificate) setup — Teams Attendance extractor

Enables `Bulk-ExtractTeamsAttendance-Phased.ps1` **mode [2]** to read attendance for **all**
targeted organizers, unattended. One-time setup. Requires **Global Administrator** (for admin
consent + the Teams application access policy).

Total time: ~15 min of work + up to **30 min** for the Teams policy to propagate.

---

## Part A — Create the certificate (on the machine that will RUN the script)

Run in PowerShell 7 as the account that will run the script:

```powershell
$cert = New-SelfSignedCertificate -Subject "CN=RaksoTeamsAttendance" `
  -CertStoreLocation "Cert:\CurrentUser\My" `
  -KeyExportPolicy Exportable -KeySpec Signature -NotAfter (Get-Date).AddYears(2)

$cert.Thumbprint                                   # <-- you enter THIS in the script (mode 2)
Export-Certificate -Cert $cert -FilePath "$HOME\RaksoTeamsAttendance.cer" | Out-Null
```

- The **private key stays in this machine's** `Cert:\CurrentUser\My`. The script finds it by thumbprint.
- You upload only the **public** `.cer` to Entra.
- If a *different* machine will run it, create the cert there instead (or export the `.pfx` with private key and import it on that machine).

---

## Part B — Register the app (Entra)

1. Go to **entra.microsoft.com** → **Identity → Applications → App registrations → New registration**.
2. Name: `Rakso Teams Attendance Export`. Account types: **Single tenant**. Click **Register**.
3. On **Overview**, copy:
   - **Application (client) ID**  → you enter this in the script
   - **Directory (tenant) ID**    → you enter this in the script

---

## Part C — Upload the certificate

App → **Certificates & secrets → Certificates → Upload certificate** → select `RaksoTeamsAttendance.cer` → **Add**.
Confirm the thumbprint shown matches `$cert.Thumbprint` from Part A.

---

## Part D — Add Graph application permissions + admin consent

App → **API permissions → Add a permission → Microsoft Graph → Application permissions**, add all four:

| Permission | Why |
|---|---|
| `Calendars.Read` | read each organizer's calendar to find their Teams meetings |
| `OnlineMeetings.Read.All` | resolve a meeting's JoinWebUrl → onlineMeeting id |
| `OnlineMeetingArtifact.Read.All` | read the attendance reports + records |
| `User.Read.All` | resolve organizer display names |

Then click **Grant admin consent for <tenant>** and confirm every row shows a green **Granted** check.

---

## Part E — Teams application access policy (the step most people miss)

App-only calls to a user's online meetings return **403** unless an application access policy
authorizes the app for that user. Run in PowerShell:

```powershell
Install-Module MicrosoftTeams -Scope CurrentUser -Force
Connect-MicrosoftTeams                       # sign in as an admin

New-CsApplicationAccessPolicy -Identity "RaksoTeamsAttendance" `
  -AppIds "<APPLICATION_CLIENT_ID>" -Description "Attendance export"

# Option 1 - tenant-wide (simplest; every organizer covered):
Grant-CsApplicationAccessPolicy -PolicyName "RaksoTeamsAttendance" -Global

# Option 2 - only specific organizers (run once per teacher):
# Grant-CsApplicationAccessPolicy -PolicyName "RaksoTeamsAttendance" -Identity "teacher1@k12.adamson.edu.ph"
```

> **Wait up to 30 minutes** for this to take effect before running the script.

---

## Part F — Run the script

Run `Bulk-ExtractTeamsAttendance-Phased.ps1`. In **Phase 2** choose **[2] App-only (certificate)** and enter:

- Tenant ID (GUID)
- App (client) ID
- Certificate thumbprint (from Part A)

---

## Quick manual smoke test (optional, before a full run)

```powershell
Connect-MgGraph -TenantId <TENANT_ID> -ClientId <APP_ID> -CertificateThumbprint <THUMB> -NoWelcome
(Get-MgContext).AuthType        # should be AppOnly
Get-MgUserOnlineMeeting -UserId "teacher1@k12.adamson.edu.ph" `
  -Filter "JoinWebUrl eq '<paste a Teams join URL from that teacher>'"
```

A returned meeting = the app + policy work. A 403 = policy not granted or not yet propagated.

---

## Troubleshooting

| Symptom | Cause / fix |
|---|---|
| `403 Forbidden` on onlineMeetings/attendance | Application access policy missing, not granted to that organizer, or <30 min old. Re-check Part E; wait. |
| Calendar reads OK but attendance 403 | `OnlineMeetingArtifact.Read.All` not added or admin consent not granted (Part D). |
| `Certificate ... not found` at connect | Thumbprint wrong, or cert not in the **runner account's** `Cert:\CurrentUser\My`. |
| 0 attendance rows for a meeting | Meeting had no attendance report (didn't occur, or reports disabled by Teams meeting policy), or it was a "Meet now" with no calendar event. |
| `Connect-MgGraph` cert works but Graph 401 later | Consent not granted, or the app has delegated (not **Application**) permissions. |

Security: the certificate's private key never leaves the runner machine. Rotate the cert before
`NotAfter`. Prefer per-organizer grants (Option 2) over `-Global` if you want to limit scope.
