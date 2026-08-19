# PowerShell Script Register — Latest & Final Versions
**Project:** Technical Ops – PowerShell Automation
**Updated:** 2026-08-19 · Repo: `SCRIPTS\Github files` → github.com/Hayleyllujah12/EntraID

---

## 1. Master list

| # | Script | Version | Updated | Type | Status |
|---|--------|---------|---------|------|--------|
| 1 | `Bulk-ResetPasswords-Phased.ps1` | **v3.7.0** | 2026-08-19 | Write | Final — gold standard |
| 2 | `Bulk-DeactivateUsers-Phased.ps1` | **v1.1** | 2026-08-19 | Write | Final |
| 3 | `Bulk-CreateUsers-AssignLicense-Phased.ps1` | **v2.0** | 2026-08-19 | Write | Final — 1 known gap |
| 4 | `Bulk-ExtractAuthMethods-Phased.ps1` | **v2.0** | 2026-08-19 | Read-only | Final |
| 5 | `Bulk-ExtractM365StorageReport-Phased.ps1` | **v2.0** | 2026-08-19 | Read-only | Final |

Repo scaffolding: `README.md`, `.gitignore`.

**Superseded — archive, do not run:** the 2026-06-16 copy of #3, and `Bulk Create Users +Assign A3 license student v1.txt`.

---

## 2. 2026-08-19 de-hardcoding pass

Applied across all five scripts: **no hardcoded tenant, no hardcoded SKU, no hardcoded input/output paths.**

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

## 3. Compliance

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

---

## 4. Remaining backlog

1. **CreateUsers v2.1 — split credentials out of the audit log.** Temporary passwords are still
   written into the same CSV as the audit trail. This is the last real security gap in the set.
2. **CreateUsers — add `Invoke-GraphWithRetry`, correlation ID, heartbeat, failed-rows CSV** to
   bring it level with ResetPasswords and DeactivateUsers.
3. **StorageReport — parameterize the `D30` period** (currently fixed).
4. **Archive** the two superseded copies in the project.

---

## 5. Companion HTML tools

| Tool | Updated | Purpose |
|---|---|---|
| `Entra_Comparison_Tool_v3.html` | 2026-06-25 | Latest Entra comparison tool |
| `Entra ID masterlist Comparison x New Deployment Accounts.html` | 2026-06-25 | Masterlist vs. new-deployment reconciliation |
| `(Working file) V2 soc-dashboard.html` | 2026-06-23 | SOC dashboard, working file |
| `M365_Bulk_User_Generator.html` | 2026-06-23 | CSV generator for CreateUsers |
