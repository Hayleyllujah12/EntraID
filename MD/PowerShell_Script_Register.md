# PowerShell Script Register — Latest & Final Versions
**Project:** Technical Ops – PowerShell Automation
**Updated:** 2026-09-02 · Repo: `SCRIPTS\Github files` → github.com/Hayleyllujah12/EntraID

---

## 1. Master list

| # | Script | Version | Updated | Type | Status |
|---|--------|---------|---------|------|--------|
| 1 | `Bulk-ResetPasswords-Phased.ps1` | **v3.7.0** | 2026-08-19 | Write | Final — gold standard |
| 2 | `Bulk-DeactivateUsers-Phased.ps1` | **v1.1** | 2026-08-19 | Write | Final |
| 3 | `Bulk-CreateUsers-AssignLicense-Phased.ps1` | **v2.0** | 2026-08-19 | Write | Final — 1 known gap |
| 4 | `Bulk-ExtractAuthMethods-Phased.ps1` | **v2.0** | 2026-08-19 | Read-only | Final |
| 5 | `Bulk-ExtractM365StorageReport-Phased.ps1` | **v2.0** | 2026-08-19 | Read-only | Final |
| 6 | `Bulk-SwapLicenses-Phased.ps1` | **v2.0** | 2026-09-02 | Write | Final — 🆕 new |
| 7 | `Bulk-ExtractTeamsAttendance-Phased.ps1` | **v1.3.0** | 2026-08-19 | Read-only | Final — 🆕 new |
| 8 | `Bulk-ExtractTeamsMeetingLogs-Phased.ps1` | **v1.0.1** | 2026-09-02 | Read-only | Final — 🆕 new |
| 9 | `Bulk-UpdateDisplayNames-Phased.ps1` | **5-phase** | 2026-08-19 | Write | Final — 🆕 new |
| 10 | `5 Phase - Update Job Title.txt` (Job Title / Dept) | **v1.0.0** | 2026-07-30 | Write | Final — 🆕 new |

Repo scaffolding: `README.md`, `.gitignore`.

**Superseded — archive, do not run:** the 2026-06-16 copy of #3, and `Bulk Create Users +Assign A3 license student v1.txt`.

---

## 2. 🆕 New since 2026-08-19

Five scripts already in the repo were not covered by the last register. Now folded in:

| # | Script | What it does | Notes |
|---|--------|--------------|-------|
| 6 | **`Bulk-SwapLicenses-Phased.ps1`** v2.0 | Bulk-swaps Entra ID education licences from a CSV of UPNs. Assigns one licence or a stack, and when a stack carries mutually-exclusive services it keeps the higher tier and disables the lower one automatically. | DRY RUN by default; live mode explicit. Solves the A1/A3 SharePoint *Plan 1 vs Plan 2* conflict the admin centre rejects — groups service plans into conflict families, elects a winner per family. Group-assigned licences detected & reported, never silently skipped. In active use (run logs from 2026-09-02). |
| 7 | **`Bulk-ExtractTeamsAttendance-Phased.ps1`** v1.3.0 | Full Teams meeting attendance (name, email, role, join/leave, duration) for meetings **organized** by a user list, via the Graph attendance-report API — the reliable source for complete rosters (incl. students) the UAL does not store. | Runtime auth: delegated interactive, app-only cert, or app-only secret. Times in CSV are Manila (PHT); raw JSONL keeps UTC. Least-privilege (dropped `User.Read.All` in v1.1.0). |
| 8 | **`Bulk-ExtractTeamsMeetingLogs-Phased.ps1`** v1.0.1 | Read-only extract of Teams meeting activity from the Unified Audit Log by date range: joins `MeetingDetail` + `MeetingParticipantDetail`. Two scope modes (organized-by vs own-participation). | Window sliced into ≤ChunkDays chunks, paged with `ReturnLargeSet` to stay under the 50k per-session ceiling. All timestamps UTC. v1.0.1 fixed the `Sanitize-Path/Upn` `String.Replace(char,'')` bug. |
| 9 | **`Bulk-UpdateDisplayNames-Phased.ps1`** | 5-phase copy-paste tool that updates ONLY `DisplayName` on existing users, matched by UPN. | Sanitization, retry/backoff, progress checkpoints, timestamped CSV log. Scope `User.ReadWrite.All`. |
| 10 | **`5 Phase - Update Job Title.txt`** v1.0.0 | Bulk-updates profile fields (First/Last/Display name, Job Title, Department) from Excel or CSV, writing only non-blank cells that differ from the current tenant value. | DryRun + Y/N confirm in Phase 4. Refactor of the legacy USER MANAGEMENT update script. Scope `User.ReadWrite.All`; `ImportExcel` for .xlsx input. |

---

## 3. 2026-08-19 de-hardcoding pass

Applied across the original five scripts: **no hardcoded tenant, no hardcoded SKU, no hardcoded input/output paths.**

| Script | v→ | What changed |
|---|---|---|
| ResetPasswords | 3.6.0 → **3.7.0** | Removed operator-specific default CSV + log folder (pointed at one admin's OneDrive). Tenant was already prompt-only. |
| DeactivateUsers | 1.0 → **1.1** | Removed `C:\Users\LITO\...` default input + log folder. Tenant and SKU picker were already compliant. |
| CreateUsers | — → **2.0** | Tenant GUID removed → runtime prompt + GUID validation. **Four hardcoded education SKU GUIDs removed** → live `Get-MgSubscribedSku` numbered picker. Paths de-hardcoded. Added `Remove-InvisibleChars` on typed paths. |
| ExtractAuthMethods | — → **2.0** | Tenant GUID removed → prompt + validation. Paths de-hardcoded. Version header added. |
| ExtractM365StorageReport | — → **2.0** | Tenant GUID removed → prompt + validation. Output folder de-hardcoded. **`CURRENT STORAGE ( JUNE)` column label was frozen at June** — now derives from the run month, overridable. Version header added. |

### Path resolution pattern (all five)

```powershell
$ScriptDir = ''
try { if ($PSScriptRoot) { $ScriptDir = $PSScriptRoot } } catch { }
if (-not $ScriptDir) {
    try {
        if ($MyInvocation.MyCommand.Path) { $ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
    } catch { }
}
if (-not $ScriptDir) { $ScriptDir = (Get-Location).Path }
```

Run as a file → resolves to the script's folder. Pasted block-by-block → falls back to the
working directory. `try/catch` wrappers keep it safe under `Set-StrictMode -Version Latest`.
Every default is still overridable at the prompt.

### Verification performed

- All 5 parse cleanly through `[System.Management.Automation.Language.Parser]::ParseFile` (PowerShell 7.4.6)
- Path resolution smoke-tested in both file-run and paste-in modes under StrictMode
- GUID validation regex tested against valid / malformed / empty input
- Repo-wide grep confirms zero remaining tenant GUIDs, SKU GUIDs, or `C:\Users\...` literals

---

## 4. Compliance (original five)

| Control | Reset | Deactivate | CreateUsers | AuthMethods | Storage |
|---|---|---|---|---|---|
| Versioned header | ✅ | ✅ | ✅ | ✅ | ✅ |
| No hardcoded Tenant ID | ✅ | ✅ | ✅ | ✅ | ✅ |
| No hardcoded SKU GUID | n/a | ✅ | ✅ | n/a | n/a |
| No hardcoded paths | ✅ | ✅ | ✅ | ✅ | ✅ |
| Unicode hardening | ✅ | ✅ | ✅ | ✅ | ✅ |
| Retry / backoff | ✅ | ✅ | ❌ | ✅ | ✅ |
| Correlation ID | ✅ | ✅ | ❌ | ❌ | ❌ |
| Heartbeat + flush | ✅ | ✅ | ❌ | ✅ | n/a |
| Failed-rows CSV | ✅ | ✅ | ❌ | ❌ | n/a |
| Secrets split from audit log | ✅ | ✅ | ❌ | n/a | n/a |
| `return` not `exit` | ✅ | ✅ | ✅ | ✅ | ✅ |

> New scripts (#6–#10) are not yet scored against this matrix — audit pass pending (see backlog).

---

## 5. Remaining backlog

1. **CreateUsers v2.1 — split credentials out of the audit log.** Temporary passwords are still
   written into the same CSV as the audit trail. This is the last real security gap in the set.
2. **CreateUsers — add `Invoke-GraphWithRetry`, correlation ID, heartbeat, failed-rows CSV** to
   bring it level with ResetPasswords and DeactivateUsers.
3. **StorageReport — parameterize the `D30` period** (currently fixed).
4. **Archive** the two superseded copies in the project.
5. **Score new scripts (#6–#10) against the §4 compliance matrix** and add rows.
6. **De-hardcoding audit** of the new scripts (tenant / paths / SKU) to confirm parity with the original five.

---

## 6. Companion HTML / doc tools

| Tool | Updated | Purpose |
|---|---|---|
| `Teams-Attendance-Dashboard.html` | 2026-08-22 | Dashboard for the Teams attendance/meeting-log extracts |
| `Teams-Attendance-Entra-Setup-Guide.docx` | 2026-08-22 | App-only (Entra) setup guide for the attendance extractor |
| `Entra_Comparison_Tool_v3.html` | 2026-06-25 | Latest Entra comparison tool |
| `Entra ID masterlist Comparison x New Deployment Accounts.html` | 2026-06-25 | Masterlist vs. new-deployment reconciliation |
| `(Working file) V2 soc-dashboard.html` | 2026-06-23 | SOC dashboard, working file |
| `M365_Bulk_User_Generator.html` | 2026-06-23 | CSV generator for CreateUsers |
