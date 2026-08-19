<<<<<<< HEAD
# EntraID PowerShell Automation

Copy-paste-into-terminal PowerShell tooling for Microsoft Entra ID / Microsoft 365
administration. Every script runs in the same 5-phase structure so it can be pasted
one block at a time and stopped at any phase boundary.

Maintained by Rakso CT Education IT.

## Scripts

| Script | Version | Type | Purpose |
|---|---|---|---|
| `Bulk-ResetPasswords-Phased.ps1` | 3.7.0 | write | Bulk password reset with CSPRNG generation, batched writes, resume, force sign-out |
| `Bulk-DeactivateUsers-Phased.ps1` | 1.1 | write | Offboarding: license downgrade, session revoke, block sign-in |
| `Bulk-CreateUsers-AssignLicense-Phased.ps1` | 2.0 | write | Bulk user creation + license assignment |
| `Bulk-ExtractAuthMethods-Phased.ps1` | 2.0 | read-only | Per-user registered MFA / auth methods |
| `Bulk-ExtractM365StorageReport-Phased.ps1` | 2.0 | read-only | Tenant OneDrive + SharePoint + Exchange storage snapshot |

## Requirements

- **PowerShell 7+** (the Graph SDK auth stack is unreliable on Windows PowerShell 5.1)
- Microsoft Graph PowerShell SDK — each script installs the submodules it needs, after a Y/N prompt
- `ImportExcel` — only for `Bulk-DeactivateUsers` when the input is `.xlsx`

## Conventions

These are enforced across every script in this repo:

- **No hardcoded tenant ID.** Prompted at runtime and GUID-validated before connecting.
- **No hardcoded SKU GUID.** Licenses are listed live from `Get-MgSubscribedSku` and picked by number.
- **No hardcoded paths.** Defaults derive from the script's own folder and stay overridable at the prompt.
- **No credentials in the repo.** See `.gitignore` — logs, credential files and input CSVs are all excluded.
- `return`, never `exit` — `exit` closes the operator's terminal mid-paste.
- `List[object]` accumulators, never `+=` in a loop.
- ASCII console markers (`[OK]`, `[FAIL]`, `[SKIP]`, `[WARN]`) — no emoji.

## The 5 phases

1. **Configure paths** — input file, output folder, correlation ID
2. **Connect** — tenant prompt, module check, scope consent, sign-in mode
3. **Verify input** — row count, header validation, UPN sanitization report
4. **Confirm** — summary of exactly what is about to happen, then Y/N
5. **Execute** — retry-wrapped Graph calls, heartbeat, periodic flush, audit CSV + failed-rows CSV

## Safety

Every write script supports a dry run and prints an execution summary. Read-only scripts
(`Bulk-Extract*`) never modify a user attribute.

`Bulk-ResetPasswords` writes its credentials file **before** any reset is applied, so an
interrupted run can never leave a user with a password nobody recorded. Distribute that
file, then delete it.

## Known gaps

- `Bulk-CreateUsers-AssignLicense` still writes temporary passwords into the same log CSV
  as the audit trail. Split these into a separate restricted credentials file before using
  it against a production tenant.
- `Bulk-ExtractM365StorageReport` has a fixed `D30` reporting period.
=======
[Bulk User Deployments] This script uses powershell 7 for bulk deployment of users account in Microsoft's backend EntraID
Uses csv file that contains the list of information of users per row
Required headers are needed to match the script
Auto sanitize the text like trailing spaces, leading spaces, special character like enye (ñ) into regular "n" character
[Bulk Reset PowerShell] Upload list of users in csv or excel file and bulk reset using PowerShell
>>>>>>> 23f8d2f24aa6598c56ec68bbb402d19837cbf3d2
