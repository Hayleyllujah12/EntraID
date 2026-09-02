# Bulk-SwapLicenses-Phased.ps1

**Version:** v2.0 · **Updated:** 2026-09-02 · **Type:** Write · **Status:** Final
**Project:** Technical Ops – PowerShell Automation

Swaps Entra ID education licences in bulk from a CSV of UPNs. Assigns **one licence or a
stack of several**, and when a stack carries mutually-exclusive services it keeps the
higher tier and disables the lower one automatically — the thing the admin centre refuses
to do for you.

---

## 1. What it does

| | |
|---|---|
| **Input** | CSV with `UserPrincipalName` (or `Username` / `UPN` — auto-detected). Optional `UsageLocation`. |
| **Action** | Per user: remove licences per the chosen strategy, then assign the target licence or stack in one atomic `Set-MgUserLicense` call. |
| **Output** | Timestamped audit CSV with before/after licence state per row. |
| **Default mode** | DRY RUN. Live mode is an explicit choice. |

**Does not** create, delete or disable users, touch group membership, or change mailbox
settings. Group-assigned licences are detected and reported, never silently skipped.

---

## 2. The stacking problem it solves

Office 365 A1 and A3 both ship SharePoint for Education — but at different tiers
(*Plan 1* vs *Plan 2*). They are different `ServicePlanId`s in the same mutually-exclusive
family, so the admin centre rejects the pair outright:

> You can't assign licenses that contain these conflicting services:
> SharePoint (Plan 2) for Education, SharePoint (Plan 1) for Education.

Comparing plan IDs alone never finds this — the IDs differ. The script instead groups every
service plan into a **conflict family** with a tier rank, elects one winner per family, and
disables every other member.

### Election rule

1. Highest **plan tier** in the family wins.
2. Tie → the **richer SKU** (more service plans) wins.
3. Still tied → **stack order** (the first licence you entered).

Tier is derived from the service plans a SKU actually carries, never guessed from its name —
`STANDARDWOFFPACK_FACULTY` contains no "A1" to pattern-match on, and
`ENTERPRISEPACKPLUS_STUUSEBNFT` no "A3".

### Worked example — the live tenant, selection `2,4`

```
Target STACK (2 licenses):
  BASE    ENTERPRISEPACKPLUS_STUUSEBNFT      (A3 students use benefit)
  STACKED STANDARDWOFFPACK_FACULTY           (A1 faculty)

CONFLICTING SERVICES FOUND (4) - keeping the higher tier:
  KEEP    SHAREPOINTENTERPRISE_EDU     from ENTERPRISEPACKPLUS_STUUSEBNFT
  DISABLE SHAREPOINTSTANDARD_EDU       from STANDARDWOFFPACK_FACULTY
  KEEP    EXCHANGE_S_ENTERPRISE        from ENTERPRISEPACKPLUS_STUUSEBNFT
  DISABLE EXCHANGE_S_STANDARD          from STANDARDWOFFPACK_FACULTY
  KEEP    OFFICESUBSCRIPTION           from ENTERPRISEPACKPLUS_STUUSEBNFT
  DISABLE OFFICEMOBILE_SUBSCRIPTION    from STANDARDWOFFPACK_FACULTY
  KEEP    STREAM_O365_E3               from ENTERPRISEPACKPLUS_STUUSEBNFT
  DISABLE STREAM_O365_E1               from STANDARDWOFFPACK_FACULTY

  ENTERPRISEPACKPLUS_STUUSEBNFT    17 of 17 services active
  STANDARDWOFFPACK_FACULTY          0 of 16 services active
```

A3 won all four families **despite A1 being enterable first**, confirming election is by
tier and not by position.

### Conflict families covered

SharePoint · Exchange · Office apps · Office web apps · Project · Skype/Teams · Entra ID ·
Rights Management · Stream · Power Automate · Power Apps · Power BI · Forms · To-Do ·
Yammer · Intune.

Plans outside these families fall back to same-`ServicePlanId` de-duplication, so an
unknown workload present in both SKUs is still disabled on the lower one. If Graph rejects
an assignment over a conflict the table does not know, Phase 5 retries the operation as
remove-first-then-add.

---

## 3. Phases

| Phase | Does |
|---|---|
| **1 — Paths** | CSV + log folder prompts, defaults beside the script. Strips invisible Unicode from pasted paths. Validates the CSV exists. |
| **2 — Connect** | Module check, Tenant ID prompt, sign-in mode (device code / popup / reuse), then loads the live SKU catalogue with `Get-MgSubscribedSku -All`. |
| **3 — Verify CSV** | Row count, UPN column auto-detect, 20-row preview. |
| **4 — Configure** | Removal strategy → target licence(s) → conflict resolution → run mode → capacity check → confirmation. |
| **5 — Execute** | Atomic swap per user, heartbeat every 25 rows, log flush every 100, summary dashboard. |

### Removal strategies (Phase 4 Step 1)

| | Strategy | Removes |
|---|---|---|
| **A** | Specific | One named SKU, only if the user holds it. |
| **B** | All | Everything the user holds except the target stack. |
| **C** | EDU only *(recommended)* | Education-family SKUs only — preserves Power BI Free, Flow Free, Teams add-ons. |

All three exclude **every** SKU in the target stack from removal, so a strategy can never
strip a licence you just asked it to assign.

### Selecting a stack (Phase 4 Step 2)

```
Enter # of license(s) to ASSIGN (1-5, comma-separated, or Q): 2,4
```

One number → single licence, identical to v1.0 behaviour. Several → stacked. The first is
the base. Duplicates are ignored with a note; a SKU set for removal cannot also be assigned.

After resolution you get one `Y/N` to accept it. Answering `N` assigns everything enabled —
which Graph will almost certainly reject, and the script warns you so.

---

## 4. Requirements

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser -Force
```

Uses `Microsoft.Graph.Users`, `Microsoft.Graph.Identity.DirectoryManagement`,
`Microsoft.Graph.Users.Actions`. **PowerShell 7+.**

| Scope | Why |
|---|---|
| `User.ReadWrite.All` | Read licence state and call `assignLicense` per user. |
| `Organization.Read.All` | Enumerate subscribed SKUs and their service plans. |

Admin role: **User Administrator** or **Global Administrator**.

---

## 5. CSV template

```csv
UserPrincipalName,UsageLocation
juan.delacruz@school.edu.ph,PH
maria.santos@school.edu.ph,PH
```

`UsageLocation` is optional — users without one are backfilled from the column, else from
`$DefaultUsageLocation` (`PH`). Graph rejects `assignLicense` without it.

---

## 6. Log columns

`Timestamp` · `Row` · `UserPrincipalName` · `RunMode` · `Strategy` · `OldLicenses` ·
`NewLicenses` · `LicensesAdded` · `LicensesRemoved` · **`DisabledPlans`** ·
`GroupAssignedLicenses` · `Status` · `Message`

`Status` ∈ `Success` · `WouldChange` · `NoChange` · `Failed` · `NotFound` · `Skipped`

---

## 7. Verification

```powershell
# Confirm the stack landed with the right services active
Get-MgUserLicenseDetail -UserId <upn> |
  Select-Object SkuPartNumber, @{n='Active';e={
    ($_.ServicePlans | Where-Object ProvisioningStatus -eq 'Success').ServicePlanName -join ','}}

# Tenant-level seat counts after a batch
Get-MgSubscribedSku -All |
  Select-Object SkuPartNumber, ConsumedUnits, @{n='Enabled';e={$_.PrepaidUnits.Enabled}}
```

Expect the lower-tier SKU in a stack to show an empty or near-empty `Active` list. That is
correct, and it is also the signal described in §9.

---

## 8. Rollback

The log's `OldLicenses` column holds each user's pre-run state. To reverse one user:

```powershell
Set-MgUserLicense -UserId <upn> `
  -AddLicenses @(@{SkuId='<originalSkuId>'; DisabledPlans=@()}) `
  -RemoveLicenses @('<newSkuId>')
```

To reverse a whole run, build a CSV from the log's failed/changed rows and re-run with the
original licence as the target.

> **Caution:** removing a SKU that supplied OneDrive or SharePoint starts the **30-day
> retention clock** on that user's OneDrive if no remaining SKU provides it. This is
> precisely why the resolver keeps the workload on the higher tier rather than dropping it.

---

## 9. Known consideration — the zero-service stack

In the A3 + A1 combination the A1 licence ends with **0 of 16 services active**. Every one
of its workloads is either a lower tier or a duplicate of A3's, so it delivers nothing to
the user while still consuming a seat.

That is mathematically correct and it is what makes the stack assignable. But unless the
lower licence is held deliberately — to reserve a seat count, satisfy a group-licensing
rule, or stage a later downgrade — assigning A3 alone produces an identical user experience
for one fewer licence. On the live tenant `STANDARDWOFFPACK_FACULTY` has only **75 free
units**, so a stacked run is capped at 75 users regardless of the 449 A3 seats available.

Phase 4 flags per-SKU capacity before you confirm.

---

## 10. Compliance

| Control | Status |
|---|---|
| Versioned header | ✅ v2.0 |
| No hardcoded Tenant ID | ✅ prompt + validation |
| No hardcoded SKU GUID | ✅ live `Get-MgSubscribedSku -All` picker |
| No hardcoded paths | ✅ `$ScriptDir` pattern (v2.0) |
| Unicode hardening | ✅ paths + UPNs, NFD-normalised |
| Retry / backoff | ✅ `Invoke-GraphWithRetry`, honours `Retry-After` |
| Heartbeat + flush | ✅ 25 / 100 rows |
| DryRun default | ✅ |
| `return` not `exit` | ✅ |
| StrictMode-safe | ✅ verified under `Set-StrictMode -Version Latest` |
| Correlation ID | ❌ backlog |
| Failed-rows CSV | ❌ backlog |
| Secrets split from audit log | n/a — touches no credentials |

---

## 11. Version history

### v2.0 — 2026-09-02

**Stacked licences.** Phase 4 Step 2 accepts several menu numbers; the payload carries one
entry per SKU with its own `DisabledPlans`.

**Automatic conflict resolution.** New service-plan family table with tier ranks; highest
tier in each family is kept, all lower and duplicate members disabled. Fixes the admin
centre's *"conflicting services"* rejection for A1 + A3.

**SKU tier derived, not guessed.** Read from the service plans a SKU carries rather than
its part number, which carries no reliable tier marker.

**Stack-aware removal.** All three strategies now exclude every target SKU
(`$TargetSkuIds -notcontains $_`), not just one. Previously strategy B or C could strip a
SKU that was about to be assigned.

**Plan-drift detection.** A user already holding the target SKUs but with the wrong disabled
plans is re-applied rather than logged `NoChange` — re-adding a SKU is how Graph updates its
disabled plans.

**Conflict self-heal.** If Graph refuses a combined add+remove as conflicting, the operation
retries as remove-first-then-add.

**Per-SKU capacity check.** A stacked assignment consumes one unit of each SKU, so each is
checked against the row count separately.

**`Get-MgSubscribedSku -All`.** Without `-All` a tenant with many SKUs can page-truncate; a
missing SKU would break family resolution silently.

**De-hardcoded paths.** Operator-specific `C:\Users\LITO\Downloads\...` defaults replaced
with the register's `$ScriptDir` pattern. Brings the script in line with the 2026-08-19 pass.

**Version header added.**

**New log column** `DisabledPlans`.

#### Fixed in v2.0 development

| Symptom | Cause | Fix |
|---|---|---|
| `The property 'SkuId' cannot be found on this object` at the licence-selection prompt | `$picked.SkuId` member enumeration on an empty array throws under `Set-StrictMode -Version Latest`, which persists in a session from any previously pasted script | Track selections in a typed `List[string]` and test with `.Contains()`; two `+=` array builds converted to lists |
| Stack still rejected by Graph as conflicting | Resolver compared only `ServicePlanId`, so cross-tier conflicts (SharePoint Plan 1 vs Plan 2 — different IDs) were never disabled | Conflict-family table with tier ranks |

### v1.0 — baseline

Single-licence swap: remove existing licence(s) per strategy A/B/C, assign exactly one
target, guarantee no accumulation. Dry run, group-assignment detection, UPN sanitisation,
retry/backoff, heartbeat, periodic flush, before/after audit CSV.

---

## 12. Backlog

1. **Correlation ID** — add a per-run GUID stamped on every log row, to match
   `Bulk-ResetPasswords-Phased.ps1`.
2. **Failed-rows CSV** — emit `*_FAILED.csv` so only failures need re-running.
3. **Group-licensing awareness** — a licence inherited from a group cannot be removed via
   the user API. Currently detected and logged; could pre-flight `licenseAssignmentStates`
   and warn in Phase 4 before any writes.
4. **Family table maintenance** — Microsoft occasionally renames or re-tiers service plans.
   The self-heal path covers gaps, but the table should be reviewed when new SKUs are bought.
5. **Reverse-run helper** — generate a rollback CSV directly from a completed log.
