<#
============================================================
 Bulk-DeactivateUsers-Phased.ps1  (v1.1)
============================================================
.SYNOPSIS       Bulk-deactivate (offboard) Entra ID / M365 users from an Excel or CSV list.
.DESCRIPTION    For each user in the input file this script will:
                  1. (optional) Reset the password to a random strong value (NOT read from the file).
                  2. Downgrade the license to a single chosen SKU (A1 for Students by default) -
                     removing every other assigned license.
                  3. Revoke all active sign-in sessions / refresh tokens.
                  4. Block sign-in (AccountEnabled = $false).
                It does NOT delete the mailbox, remove the user object, or empty OneDrive.
                Supports a DryRun preview with a live-apply fallback (no reconnect / re-scan).
                Privileged accounts (members of any active directory role) are skipped by a guard.
.AUTHOR         Generated with Claude for Rakso CT Education IT.
.VERSION        1.1
.DATE           2026-08-19
.REQUIREMENTS   PowerShell 7+; Microsoft.Graph SDK submodules listed in Phase 2; ImportExcel (Excel input only).
.PERMISSIONS    User.ReadWrite.All, Organization.Read.All, RoleManagement.Read.Directory
.SAFETY         DryRun preview + Y/N confirmation before any write. Privileged-account guard on by default.
.CHANGELOG      v1.0 - Initial release.
                v1.1 - Removed the hardcoded default input path and log folder
                       (C:\Users\LITO\...). Both defaults are now derived at runtime
                       from the script's own folder and stay overridable at the prompt,
                       so nothing machine-specific is stored in the file. Tenant ID was
                       already prompt-only and the license SKU was already chosen from
                       a live Get-MgSubscribedSku picker; both are unchanged.
============================================================
#>
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ============================================================
# HELPERS  (paste-once at top of file)
# ============================================================

function Repair-PastedPath {
    # Strips invisible/formatting chars + a single wrapping quote pair (handles Ctrl+Shift+C "Copy as path").
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
    $p = $Path.Trim().Trim('"').Trim("'").Trim()
    $invisible = @([char]0x00A0,[char]0x200B,[char]0x200C,[char]0x200D,[char]0x200E,[char]0x200F,
                   [char]0x202A,[char]0x202B,[char]0x202C,[char]0x202D,[char]0x202E,
                   [char]0x2060,[char]0xFEFF)
    foreach ($c in $invisible) { $p = $p.Replace([string]$c,'') }   # [string] cast: avoids Replace(char,char) overload bug
    return $p.Trim()
}

function Sanitize-Upn {
    # Normalizes/cleans an identity field; reports whether anything changed.
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return @{ Clean=''; Changed=$true; Reason='Empty or whitespace-only' } }
    $raw = $Value
    $nfd = $raw.Normalize([Text.NormalizationForm]::FormD)
    $sb  = [Text.StringBuilder]::new()
    foreach ($ch in $nfd.ToCharArray()) {
        if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$sb.Append($ch)
        }
    }
    $s = $sb.ToString()
    $invisible = @(
        [char]0x00A0,[char]0x1680,[char]0x2000,[char]0x2001,[char]0x2002,[char]0x2003,
        [char]0x2004,[char]0x2005,[char]0x2006,[char]0x2007,[char]0x2008,[char]0x2009,
        [char]0x200A,[char]0x200B,[char]0x200C,[char]0x200D,[char]0x200E,[char]0x200F,
        [char]0x202A,[char]0x202B,[char]0x202C,[char]0x202D,[char]0x202E,
        [char]0x202F,[char]0x205F,[char]0x2060,[char]0x3000,[char]0xFEFF
    )
    foreach ($c in $invisible) { $s = $s.Replace([string]$c,'') }    # [string] cast: avoids Replace(char,char) overload bug
    $s = $s -replace '\p{C}',''
    $s = $s.Trim().ToLowerInvariant()
    $changed = ($s -ne $raw)
    if ($s -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') { return @{ Clean=$s; Changed=$changed; Reason='Invalid UPN format' } }
    if ($s.Length -gt 113)                         { return @{ Clean=$s; Changed=$changed; Reason='Exceeds 113-char Entra limit' } }
    return @{ Clean=$s; Changed=$changed; Reason=$(if($changed){'Sanitized'}else{'Unchanged'}) }
}

function Ensure-Module {
    param([Parameter(Mandatory)][string]$Name, [string]$MinVersion)
    $existing = Get-Module -ListAvailable -Name $Name | Sort-Object Version -Descending | Select-Object -First 1
    if ($existing) {
        if ($MinVersion -and ($existing.Version -lt [version]$MinVersion)) {
            Write-Host "[WARN] $Name $($existing.Version) < required $MinVersion." -ForegroundColor Yellow
            $ans = Read-Host "Uninstall all versions and reinstall clean? (Y/N)"
            if ($ans -notmatch '^(y|yes)$') { return $false }
            try {
                Get-InstalledModule "$Name*" -AllVersions -ErrorAction SilentlyContinue | Uninstall-Module -Force -ErrorAction SilentlyContinue
                Install-Module $Name -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
            } catch { Write-Host "[FAIL] Reinstall of $Name failed: $($_.Exception.Message)" -ForegroundColor Red; return $false }
        }
        Import-Module $Name -ErrorAction Stop
        return $true
    }
    Write-Host "[MISSING] Module $Name is not installed." -ForegroundColor Yellow
    $ans = Read-Host "Install $Name from PSGallery now? (Y/N)"
    if ($ans -notmatch '^(y|yes)$') { return $false }
    try {
        Install-Module $Name -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Import-Module $Name -ErrorAction Stop
        return $true
    } catch { Write-Host "[FAIL] Install of $Name failed: $($_.Exception.Message)" -ForegroundColor Red; return $false }
}

function Invoke-GraphWithRetry {
    param([Parameter(Mandatory)][scriptblock]$Action,[int]$MaxAttempts=5,[int]$BaseDelaySeconds=2)
    for ($i=1; $i -le $MaxAttempts; $i++) {
        try { return & $Action }
        catch {
            $ex = $_.Exception
            $status = $null; $retryAfter = $null

            # StrictMode-safe: probe properties via PSObject before touching them.
            # Graph SDK throws several exception shapes; not all expose .Response.
            $resp = $null
            if ($ex -and $ex.PSObject.Properties['Response']) { $resp = $ex.Response }
            if ($resp) {
                if ($resp.PSObject.Properties['StatusCode']) {
                    try { $status = [int]$resp.StatusCode } catch { $status = $null }
                }
                if ($resp.PSObject.Properties['Headers'] -and $resp.Headers) {
                    try {
                        $ra = $resp.Headers | Where-Object { $_.Key -eq 'Retry-After' } | Select-Object -First 1
                        if ($ra) { $retryAfter = ($ra.Value | Select-Object -First 1) }
                    } catch { }
                }
            }
            if ($null -eq $status -and $ex -and $ex.PSObject.Properties['HttpStatusCode']) {
                try { $status = [int]$ex.HttpStatusCode } catch { }
            }
            # Last resort: pull a status number out of the message text.
            if ($null -eq $status -and $ex -and $ex.Message -match '\b(429|500|502|503|504)\b') {
                $status = [int]$Matches[1]
            }

            if ($status -notin 429,500,502,503,504 -or $i -eq $MaxAttempts) { throw }
            $delay = if ($retryAfter) { [int]$retryAfter } else { [Math]::Pow($BaseDelaySeconds, $i) }
            Write-Host "[RETRY $i/$MaxAttempts] HTTP $status, sleeping ${delay}s..." -ForegroundColor DarkYellow
            Start-Sleep -Seconds $delay
        }
    }
}

function Write-Log {
    param([string]$Message,[ValidateSet('INFO','WARN','ERROR','SUCCESS','AUDIT','SECURITY')][string]$Level='INFO')
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
    $line = "$ts [$Level] [$Global:CorrelationId] $Message"
    $color = @{INFO='Gray';WARN='Yellow';ERROR='Red';SUCCESS='Green';AUDIT='Cyan';SECURITY='Magenta'}[$Level]
    Write-Host $line -ForegroundColor $color
    if ($Global:JsonLogFile) {
        $obj = [ordered]@{ ts=$ts; level=$Level; correlationId=$Global:CorrelationId; message=$Message }
        ($obj | ConvertTo-Json -Compress) | Add-Content -LiteralPath $Global:JsonLogFile -Encoding UTF8
    }
}

function Import-InputFile {
    # Routes by extension: ImportExcel for .xlsx/.xls/.xlsm (no Office/COM), delimiter-sniffed Import-Csv for text.
    param([Parameter(Mandatory)][string]$Path)
    $ext = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
    if ($ext -in '.xlsx','.xls','.xlsm') {
        if (-not (Ensure-Module -Name 'ImportExcel')) { throw "ImportExcel module required to read '$ext' files." }
        return @(Import-Excel -Path $Path)
    }
    # Text file: sniff delimiter from the header line (comma vs tab vs semicolon).
    $firstLine = (Get-Content -LiteralPath $Path -TotalCount 1 -Encoding UTF8)
    $delim = ','
    if     ($firstLine -match "`t") { $delim = "`t" }
    elseif ($firstLine -match ';')  { $delim = ';'  }
    return @(Import-Csv -LiteralPath $Path -Delimiter $delim -Encoding UTF8)
}

function New-RandomPassword {
    # 16-char password meeting Entra complexity (upper/lower/digit/symbol).
    param([int]$Length = 16)
    $U='ABCDEFGHJKLMNPQRSTUVWXYZ'; $L='abcdefghijkmnpqrstuvwxyz'; $D='23456789'; $S='!@#$%^&*-_=+?'
    $all = ($U+$L+$D+$S).ToCharArray()
    $chars = @($U[(Get-Random -Max $U.Length)], $L[(Get-Random -Max $L.Length)],
               $D[(Get-Random -Max $D.Length)], $S[(Get-Random -Max $S.Length)])
    for ($i=$chars.Count; $i -lt $Length; $i++) { $chars += $all[(Get-Random -Max $all.Length)] }
    -join ($chars | Sort-Object { Get-Random })
}

function Get-GraphErrorMessage {
    # Extracts the real Graph error (code + message) from an ErrorRecord.
    # Invoke-MgGraphRequest surfaces only "Response status code ... BadRequest" in
    # .Exception.Message; the useful body is in .ErrorDetails.Message as JSON.
    param($ErrorRecord)
    $detail = $null
    if ($ErrorRecord -and $ErrorRecord.PSObject.Properties['ErrorDetails'] -and
        $ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) {
        $detail = $ErrorRecord.ErrorDetails.Message
    }
    if ($detail) {
        try {
            $obj = $detail | ConvertFrom-Json -ErrorAction Stop
            if ($obj.PSObject.Properties['error'] -and $obj.error) {
                $code = if ($obj.error.PSObject.Properties['code'])    { $obj.error.code }    else { '' }
                $msg  = if ($obj.error.PSObject.Properties['message']) { $obj.error.message } else { $detail }
                return ("[{0}] {1}" -f $code, $msg).Trim()
            }
        } catch { }
        return $detail
    }
    if ($ErrorRecord -and $ErrorRecord.Exception) { return $ErrorRecord.Exception.Message }
    return "$ErrorRecord"
}

function Get-SkuIdsFromUserObject {
    # Extracts DIRECT license SkuIds from a user object regardless of shape:
    # typed .AssignedLicenses[].SkuId, untyped hashtables, or AdditionalProperties['assignedLicenses'].
    # StrictMode-safe (checks via PSObject.Properties before access). Returns lowercased, validated GUID strings.
    param($UserObj)
    $out = [System.Collections.Generic.List[string]]::new()
    if (-not $UserObj) { return @() }
    $al = $null
    if ($UserObj.PSObject.Properties['AssignedLicenses'] -and $UserObj.AssignedLicenses) {
        $al = $UserObj.AssignedLicenses
    }
    if (-not $al -and $UserObj.PSObject.Properties['AdditionalProperties'] -and $UserObj.AdditionalProperties) {
        try { if ($UserObj.AdditionalProperties.ContainsKey('assignedLicenses')) { $al = $UserObj.AdditionalProperties['assignedLicenses'] } } catch { }
    }
    if (-not $al) { return @() }
    foreach ($item in @($al)) {
        if (-not $item) { continue }
        $sku = $null
        if ($item -is [System.Collections.IDictionary]) {
            if ($item.Contains('skuId')) { $sku = $item['skuId'] } elseif ($item.Contains('SkuId')) { $sku = $item['SkuId'] }
        } elseif ($item.PSObject.Properties['SkuId']) { $sku = $item.SkuId }
          elseif ($item.PSObject.Properties['skuId']) { $sku = $item.skuId }
        if ($sku) {
            $s = ([string]$sku).Trim().ToLowerInvariant()
            if ($s -match '^[0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12}$') { $out.Add($s) }
        }
    }
    return @($out)
}

function Get-CurrentSkuIds {
    # Authoritative, live read of a single user's DIRECT license SkuIds (lowercased, validated).
    param([Parameter(Mandatory)][string]$UserId)
    $gu = Invoke-GraphWithRetry { Get-MgUser -UserId $UserId -Property assignedLicenses -ErrorAction Stop }
    return Get-SkuIdsFromUserObject $gu
}

function Set-UserLicenseRaw {
    # Calls the assignLicense endpoint with a hand-built body so 'addLicenses' is
    # ALWAYS present (as []). Set-MgUserLicense drops the key when -AddLicenses is
    # empty, causing: Request_BadRequest ... missing parameters: addLicenses.
    param(
        [Parameter(Mandatory)][string]$UserId,
        [string[]]$AddSkuIds    = @(),
        [string[]]$RemoveSkuIds = @()
    )
    $guidRe = '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$'
    $AddSkuIds    = @($AddSkuIds    | ForEach-Object { if ($_) { ([string]$_).Trim() } } | Where-Object { $_ -match $guidRe })
    $RemoveSkuIds = @($RemoveSkuIds | ForEach-Object { if ($_) { ([string]$_).Trim() } } | Where-Object { $_ -match $guidRe })
    if (@($AddSkuIds).Count -eq 0 -and @($RemoveSkuIds).Count -eq 0) { return }   # nothing valid to send
    $addItems    = @($AddSkuIds    | ForEach-Object { '{{"disabledPlans":[],"skuId":"{0}"}}' -f $_ })
    $removeItems = @($RemoveSkuIds | ForEach-Object { '"{0}"' -f $_ })
    $json = '{"addLicenses":[' + ($addItems -join ',') + '],"removeLicenses":[' + ($removeItems -join ',') + ']}'
    Invoke-MgGraphRequest -Method POST `
        -Uri "https://graph.microsoft.com/v1.0/users/$UserId/assignLicense" `
        -Body $json -ContentType 'application/json' -ErrorAction Stop
}

# ============================================================
# PHASE 1 - Configure Paths
# ============================================================
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " PHASE 1: Configure Paths" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

# --- Dynamic path defaults ---------------------------------------------------
# No machine-specific path is stored in this file. Defaults resolve at runtime to
# the script's own folder. When the script is pasted block-by-block into a
# terminal $PSScriptRoot is empty, so the current working directory is used.
# Every default is still overridable at the prompt.
$ScriptDir = ''
try { if ($PSScriptRoot) { $ScriptDir = $PSScriptRoot } } catch { }
if (-not $ScriptDir) {
    try {
        if ($MyInvocation.MyCommand.Path) { $ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
    } catch { }
}
if (-not $ScriptDir) { $ScriptDir = (Get-Location).Path }

$DefaultInput     = Join-Path $ScriptDir 'deactivate.csv'
$DefaultLogFolder = Join-Path $ScriptDir 'Reports'

$typedInput = Read-Host "Enter input CSV/Excel path (Enter for default: $DefaultInput)"
$InputPath  = if ([string]::IsNullOrWhiteSpace($typedInput)) { $DefaultInput } else { Repair-PastedPath $typedInput }

$typedLog  = Read-Host "Enter log folder (Enter for default: $DefaultLogFolder)"
$LogFolder = if ([string]::IsNullOrWhiteSpace($typedLog)) { $DefaultLogFolder } else { Repair-PastedPath $typedLog }

if (-not (Test-Path -LiteralPath $InputPath)) { Write-Host "[FAIL] Input file not found: $InputPath" -ForegroundColor Red; return }
if (-not (Test-Path -LiteralPath $LogFolder)) { New-Item -Path $LogFolder -ItemType Directory -Force | Out-Null }

$Global:CorrelationId = [guid]::NewGuid().ToString('N').Substring(0,12)
$stamp      = Get-Date -Format 'yyyyMMdd_HHmmss'
$ScriptName = 'Bulk-DeactivateUsers'
$Global:LogFile     = Join-Path $LogFolder "${ScriptName}_${stamp}.csv"
$Global:FailedFile  = Join-Path $LogFolder "${ScriptName}_${stamp}_FAILED.csv"
$Global:JsonLogFile = Join-Path $LogFolder "${ScriptName}_${stamp}.jsonl"
$Global:PwFile      = Join-Path $LogFolder "${ScriptName}_${stamp}_RESET_PASSWORDS.csv"

Write-Host "[OK] Correlation ID : $Global:CorrelationId" -ForegroundColor Green
Write-Host "[OK] Input          : $InputPath" -ForegroundColor Green
Write-Host "[OK] Log CSV        : $Global:LogFile" -ForegroundColor Green

# ============================================================
# PHASE 2 - Connect to Microsoft Graph + Select License
# ============================================================
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " PHASE 2: Connect to Microsoft Graph" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

$Modules = @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Users',
    'Microsoft.Graph.Users.Actions',
    'Microsoft.Graph.Identity.DirectoryManagement'
)
foreach ($m in $Modules) {
    if (-not (Ensure-Module -Name $m)) { Write-Host "[FAIL] Cannot proceed without $m." -ForegroundColor Red; return }
}

$TenantId = (Read-Host "Enter Tenant ID (GUID)").Trim()
if ($TenantId -notmatch '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$') {
    Write-Host "[FAIL] Not a valid GUID." -ForegroundColor Red; return
}

$Scopes = @('User.ReadWrite.All','Organization.Read.All','RoleManagement.Read.Directory')
Write-Host ""
Write-Host "Requesting scopes:" -ForegroundColor Yellow
$Scopes | ForEach-Object { Write-Host "  - $_" }

Write-Host ""
Write-Host "Sign-in mode:" -ForegroundColor Yellow
Write-Host "  [1] Browser / DeviceCode  (default)"
Write-Host "  [2] Interactive popup"
Write-Host "  [3] Reuse existing session"
$mode = Read-Host "Select 1/2/3 (Enter=1)"
if ([string]::IsNullOrWhiteSpace($mode)) { $mode = '1' }

try {
    switch ($mode) {
        '1' { Connect-MgGraph -TenantId $TenantId -Scopes $Scopes -UseDeviceCode -NoWelcome | Out-Null }
        '2' { Connect-MgGraph -TenantId $TenantId -Scopes $Scopes -NoWelcome | Out-Null }
        '3' { if (-not (Get-MgContext)) { throw "No existing session; pick 1 or 2." } }
        default { Write-Host "[FAIL] Invalid choice." -ForegroundColor Red; return }
    }
} catch { Write-Host "[FAIL] Auth error: $($_.Exception.Message)" -ForegroundColor Red; return }

$ctx = Get-MgContext
if (-not $ctx -or -not $ctx.TenantId) { Write-Host "[FAIL] No Graph context established." -ForegroundColor Red; return }
Write-Host "[OK] Tenant : $($ctx.TenantId)" -ForegroundColor Green
Write-Host "[OK] Account: $($ctx.Account)"   -ForegroundColor Green

# --- License handling mode ---
Write-Host ""
Write-Host "License handling during deactivation:" -ForegroundColor Yellow
Write-Host "  [1] Downgrade to a specific SKU (e.g. A1 for Students) - removes all others"
Write-Host "  [2] Unlicense - remove ALL assigned licenses, assign nothing"
Write-Host "  [3] Skip - leave licenses unchanged"
$licMode = Read-Host "Select 1/2/3 (Enter=1)"
if ([string]::IsNullOrWhiteSpace($licMode)) { $licMode = '1' }
if ($licMode -notin '1','2','3') { Write-Host "[FAIL] Invalid license mode." -ForegroundColor Red; return }

$TargetSku = $null; $TargetSkuId = $null
if ($licMode -eq '1') {
    # Interactive SKU picker (never hardcode the GUID)
    Write-Host ""
    Write-Host "Fetching available SKUs from tenant..." -ForegroundColor Cyan
    $skus = Invoke-GraphWithRetry { Get-MgSubscribedSku -All } | Sort-Object SkuPartNumber
    for ($i=0; $i -lt $skus.Count; $i++) {
        $s = $skus[$i]
        Write-Host ("  [{0,2}] {1,-45} {2}/{3} consumed  ({4})" -f `
            ($i+1), $s.SkuPartNumber, $s.ConsumedUnits, $s.PrepaidUnits.Enabled, $s.SkuId) -ForegroundColor Gray
    }
    Write-Host "Pick the SKU users will be DOWNGRADED to (A1 for Students)." -ForegroundColor Yellow
    $pick = Read-Host "SKU number"
    $idx  = [int]$pick - 1
    if ($idx -lt 0 -or $idx -ge $skus.Count) { Write-Host "[FAIL] Invalid selection." -ForegroundColor Red; return }
    $TargetSku   = $skus[$idx]
    $TargetSkuId = $TargetSku.SkuId
    Write-Host "[OK] Downgrade target: $($TargetSku.SkuPartNumber) ($TargetSkuId)" -ForegroundColor Green
} elseif ($licMode -eq '2') {
    Write-Host "[OK] License mode: UNLICENSE (remove all)" -ForegroundColor Green
} else {
    Write-Host "[OK] License mode: SKIP (leave licenses unchanged)" -ForegroundColor Green
}

# ============================================================
# PHASE 3 - Verify Input Contents
# ============================================================
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " PHASE 3: Verify Input" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

$raw = Import-InputFile -Path $InputPath
Write-Host "[OK] Rows loaded (raw): $($raw.Count)" -ForegroundColor Green
if ($raw.Count -eq 0) { Write-Host "[FAIL] No rows in input." -ForegroundColor Red; return }

# Resolve the UPN column (support common header variants).
$cols   = $raw[0].PSObject.Properties.Name
$UpnCol = @('UserPrincipalName','UPN','Email','EmailAddress','Mail') | Where-Object { $_ -in $cols } | Select-Object -First 1
if (-not $UpnCol) {
    Write-Host "[FAIL] No UPN column found. Expected one of: UserPrincipalName / UPN / Email." -ForegroundColor Red
    Write-Host "       Columns present: $($cols -join ', ')" -ForegroundColor Yellow
    return
}
Write-Host "[OK] Using UPN column: $UpnCol" -ForegroundColor Green

# Drop fully-blank / no-UPN rows (Excel/CSV exports carry trailing empties).
$rows = @($raw | Where-Object { -not [string]::IsNullOrWhiteSpace($_.$UpnCol) })
$dropped = $raw.Count - $rows.Count
if ($dropped -gt 0) { Write-Host "[INFO] Dropped $dropped blank/no-UPN row(s)." -ForegroundColor DarkGray }
Write-Host "[OK] Rows to process: $($rows.Count)" -ForegroundColor Green

# Report sanitization changes (raw vs cleaned) without silently mutating.
$sanChanged = 0
foreach ($r in $rows) {
    $res = Sanitize-Upn -Value ([string]$r.$UpnCol)
    if ($res.Changed -and $res.Reason -ne 'Empty or whitespace-only') {
        $sanChanged++
        if ($sanChanged -le 10) {
            Write-Host ("[CLEAN] '{0}' -> '{1}' ({2})" -f $r.$UpnCol, $res.Clean, $res.Reason) -ForegroundColor DarkYellow
        }
    }
}
if ($sanChanged -gt 10) { Write-Host "[CLEAN] ... and $($sanChanged-10) more cleaned." -ForegroundColor DarkYellow }

Write-Host ""
Write-Host "Preview (first 3):" -ForegroundColor Yellow
$rows | Select-Object -First 3 -Property $UpnCol | Format-Table -AutoSize

# --- Prefetch the directory ONCE into a hashtable keyed by lowercased UPN ---
Write-Host "Prefetching directory (id, accountEnabled, usageLocation, assignedLicenses)..." -ForegroundColor Cyan
$allUsers = Invoke-GraphWithRetry {
    Get-MgUser -All -Property id,userPrincipalName,accountEnabled,usageLocation,assignedLicenses -ConsistencyLevel eventual
}
$dir = @{}
foreach ($u in $allUsers) { if ($u.UserPrincipalName) { $dir[$u.UserPrincipalName.ToLowerInvariant()] = $u } }
Write-Host "[OK] Directory cached: $($dir.Count) users" -ForegroundColor Green

# ============================================================
# PHASE 4 - Confirmation + Options
# ============================================================
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " PHASE 4: Confirmation" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

$DefaultUsageLocation = 'PH'
$ans = Read-Host "Reset password to a random value for each user? (Y/N, default N)"
$ResetPassword = ($ans -match '^(y|yes)$')

$ans = Read-Host "Guard: SKIP any user that holds a directory role (admins)? (Y/N, default Y)"
$GuardPrivileged = ($ans -notmatch '^(n|no)$')

$ans = Read-Host "DryRun first (preview only, no writes)? (Y/N, default Y)"
$DryRun = ($ans -notmatch '^(n|no)$')

# Build the privileged-UPN set if the guard is on.
$privUpns = New-Object 'System.Collections.Generic.HashSet[string]'
if ($GuardPrivileged) {
    Write-Host "Building privileged-account set from active directory roles..." -ForegroundColor Cyan
    $roles = Invoke-GraphWithRetry { Get-MgDirectoryRole -All }
    foreach ($role in $roles) {
        $members = Invoke-GraphWithRetry { Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -All }
        foreach ($mem in $members) {
            $mUpn = $null
            if ($mem.AdditionalProperties -and $mem.AdditionalProperties.ContainsKey('userPrincipalName')) {
                $mUpn = [string]$mem.AdditionalProperties['userPrincipalName']
            }
            if ($mUpn) { [void]$privUpns.Add($mUpn.ToLowerInvariant()) }
        }
    }
    Write-Host "[OK] Privileged accounts flagged: $($privUpns.Count)" -ForegroundColor Green
}

Write-Host ""
Write-Host "About to DEACTIVATE $($rows.Count) user(s) in tenant $($ctx.TenantId):" -ForegroundColor Yellow
$licDesc = switch ($licMode) { '1' {"Downgrade to $($TargetSku.SkuPartNumber)"} '2' {'Unlicense (remove ALL)'} '3' {'Leave unchanged'} }
Write-Host "  - License              : $licDesc"
Write-Host "  - Reset password       : $ResetPassword"
Write-Host "  - Revoke sessions      : True"
Write-Host "  - Block sign-in        : True"
Write-Host "  - Privileged guard     : $GuardPrivileged"
Write-Host "  - Mode                 : $(if($DryRun){'DRY-RUN (preview)'}else{'LIVE'})"
$go = Read-Host "Proceed? (Y/N)"
if ($go -notmatch '^(y|yes)$') { Write-Host "[SKIP] Aborted by user." -ForegroundColor Yellow; return }

# ============================================================
# PHASE 5 - Execute (Preview -> optional Live apply) + Log
# ============================================================
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " PHASE 5: Execute" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

# ---------- PASS 1: build the plan (no writes) ----------
$plan   = [System.Collections.Generic.List[object]]::new()
$skip=0
for ($i=0; $i -lt $rows.Count; $i++) {
    $rowNum = $i + 1
    $upnRes = Sanitize-Upn -Value ([string]$rows[$i].$UpnCol)
    $upn    = $upnRes.Clean
    $key    = $upn.ToLowerInvariant()

    if ($upnRes.Reason -eq 'Invalid UPN format' -or [string]::IsNullOrWhiteSpace($upn)) {
        Write-Log "Row $rowNum '$($rows[$i].$UpnCol)' - invalid UPN, skipped." 'WARN'; $skip++; continue
    }
    if (-not $dir.ContainsKey($key)) {
        Write-Log "Row $rowNum '$upn' - not found in tenant, skipped." 'WARN'; $skip++; continue
    }
    if ($GuardPrivileged -and $privUpns.Contains($key)) {
        Write-Log "Row $rowNum '$upn' - PRIVILEGED account, skipped by guard." 'SECURITY'; $skip++; continue
    }

    $u        = $dir[$key]
    # Shape-tolerant + validated; wrapped so it is always a clean array (never null/scalar).
    $curSkus  = @(Get-SkuIdsFromUserObject $u | Where-Object { $_ })
    $targetLc = if ($TargetSkuId) { ([string]$TargetSkuId).ToLowerInvariant() } else { $null }
    switch ($licMode) {
        '1' {   # Downgrade to target SKU
            $removeList = @($curSkus | Where-Object { $_ -ne $targetLc })
            $needAdd    = ($targetLc -notin $curSkus)
            $needLicense= ($needAdd -or $removeList.Count -gt 0)
        }
        '2' {   # Unlicense - remove all
            $removeList = @($curSkus)
            $needAdd    = $false
            $needLicense= ($curSkus.Count -gt 0)
        }
        '3' {   # Skip - leave licenses unchanged
            $removeList = @()
            $needAdd    = $false
            $needLicense= $false
        }
    }
    $needUsageLoc = ($needLicense -and $needAdd -and [string]::IsNullOrWhiteSpace($u.UsageLocation))
    $needBlock  = ($u.AccountEnabled -eq $true)

    $actions = @()
    if ($ResetPassword) { $actions += 'ResetPassword' }
    if ($needUsageLoc)  { $actions += 'SetUsageLocation' }
    if ($needLicense)   { $actions += 'License' }
    $actions += 'RevokeSessions'          # always part of deactivation
    if ($needBlock)     { $actions += 'BlockSignIn' }

    # If already fully deactivated and nothing to change, skip.
    if (-not $needBlock -and -not $needLicense -and -not $ResetPassword) {
        Write-Log "Row $rowNum '$upn' - already deactivated, no license change needed, skipped." 'INFO'; $skip++; continue
    }

    $plan.Add([pscustomobject]@{
        Row=$rowNum; Upn=$upn; Id=$u.Id
        RemoveList=$removeList; NeedAdd=$needAdd; NeedLicense=$needLicense
        NeedUsageLoc=$needUsageLoc; NeedBlock=$needBlock; Actions=($actions -join '+')
    })
}

Write-Host ""
Write-Host "Planned actions ($($plan.Count) user(s), $skip skipped):" -ForegroundColor Yellow
$plan | Select-Object Row, Upn, Actions | Format-Table -AutoSize

if ($plan.Count -eq 0) { Write-Host "[DONE] Nothing to apply." -ForegroundColor Cyan; return }

# ---------- Dry-run fallback: offer to apply the SAME plan live (no reconnect / re-scan) ----------
$applyLive = $true
if ($DryRun) {
    $ans = Read-Host "Apply these $($plan.Count) change(s) LIVE now? (Y/N)"
    $applyLive = ($ans -match '^(y|yes)$')
}
if (-not $applyLive) {
    Write-Host "[DRYRUN] Preview only. No changes made." -ForegroundColor Cyan
    $plan | Select-Object Row, Upn, Actions |
        Export-Csv -LiteralPath $Global:LogFile -NoTypeInformation -Encoding UTF8
    Write-Host "Preview written: $Global:LogFile" -ForegroundColor Gray
    return
}

# ---------- PASS 2: apply (writes) ----------
$results = [System.Collections.Generic.List[object]]::new()
$failed  = [System.Collections.Generic.List[object]]::new()
$pwOut   = [System.Collections.Generic.List[object]]::new()
$ok=0; $fail=0; $partial=0
$start = Get-Date
$headerWritten = $false

for ($j=0; $j -lt $plan.Count; $j++) {
    $p = $plan[$j]
    $rowNum = $p.Row
    $upn = $p.Upn; $id = $p.Id
    $done = [System.Collections.Generic.List[string]]::new()
    $errs = [System.Collections.Generic.List[string]]::new()

    # Each sub-action is independent: a failure in one does NOT abort the others.
    # Order is by security priority so access is cut even if licensing/password fails:
    #   1) Block sign-in  2) Revoke sessions  3) Reset password  4) UsageLocation  5) License
    $critFail = $false   # true if a sign-in-cutting action failed

    if ($p.NeedBlock) {
        try {
            Invoke-GraphWithRetry { Update-MgUser -UserId $id -AccountEnabled:$false -ErrorAction Stop } | Out-Null
            $done.Add('SignInBlocked')
        } catch { $errs.Add("Block: $(Get-GraphErrorMessage $_)"); $critFail = $true }
    }

    try {
        # Revoke AFTER block so no new token can be minted post-block.
        Invoke-GraphWithRetry { Revoke-MgUserSignInSession -UserId $id -ErrorAction Stop } | Out-Null
        $done.Add('SessionsRevoked')
    } catch { $errs.Add("Revoke: $(Get-GraphErrorMessage $_)"); $critFail = $true }

    if ($ResetPassword) {
        try {
            $pw = New-RandomPassword
            Invoke-GraphWithRetry {
                Update-MgUser -UserId $id -PasswordProfile @{ Password=$pw; ForceChangePasswordNextSignIn=$true } -ErrorAction Stop
            } | Out-Null
            $pwOut.Add([pscustomobject]@{ UserPrincipalName=$upn; NewPassword=$pw })
            $done.Add('PwReset')
        } catch { $errs.Add("PwReset: $(Get-GraphErrorMessage $_)") }
    }

    if ($p.NeedUsageLoc) {
        try {
            Invoke-GraphWithRetry { Update-MgUser -UserId $id -UsageLocation $DefaultUsageLocation -ErrorAction Stop } | Out-Null
            $done.Add("UsageLoc=$DefaultUsageLocation")
        } catch { $errs.Add("UsageLoc: $(Get-GraphErrorMessage $_)") }
    }

    if ($p.NeedLicense) {
        try {
            # Authoritative live re-read. Wrap in @(...|Where) so an empty/single result
            # is always a clean array (PS unrolls function array returns: empty->null, single->scalar).
            $liveSkus = @(Get-CurrentSkuIds -UserId $id | Where-Object { $_ })
            if ($licMode -eq '2') {           # Unlicense: remove everything currently held
                $rmSkus  = @($liveSkus)
                $addSkus = @()
            } else {                          # Downgrade: remove all except target, add target if missing
                $tLc     = ([string]$TargetSkuId).ToLowerInvariant()
                $rmSkus  = @($liveSkus | Where-Object { $_ -ne $tLc })
                if ($liveSkus -contains $tLc) { $addSkus = @() } else { $addSkus = @($TargetSkuId) }
            }
            $rmCount  = @($rmSkus).Count
            $addCount = @($addSkus).Count
            if ($addCount -eq 0 -and $rmCount -eq 0) {
                $done.Add('License(nochange)')
            } else {
                Invoke-GraphWithRetry { Set-UserLicenseRaw -UserId $id -AddSkuIds $addSkus -RemoveSkuIds $rmSkus } | Out-Null
                $done.Add($(if ($licMode -eq '2') { 'Unlicensed(all)' } else { "License->$($TargetSku.SkuPartNumber)" }))
            }
        } catch {
            $lm = Get-GraphErrorMessage $_
            if ($lm -match 'group|inherited') { $lm += ' (group-based license - remove the user from the licensing group instead)' }
            $errs.Add("License: $lm")
        }
    }

    # Status: FAIL if a sign-in-cutting step failed; PARTIAL if only non-critical steps failed; else OK.
    if ($critFail)          { $statusVal='FAIL' }
    elseif ($errs.Count -gt 0) { $statusVal='PARTIAL' }
    else                    { $statusVal='OK' }

    $msg = if ($errs.Count -gt 0) { $errs -join ' ; ' } else { 'Deactivated' }
    $lvl = switch ($statusVal) { 'OK' {'AUDIT'} 'PARTIAL' {'WARN'} 'FAIL' {'ERROR'} }
    Write-Log "Row $rowNum '$upn' [$statusVal] done=[$($done -join ', ')] $(if($errs.Count){"errs=[$msg]"})" $lvl

    $results.Add([pscustomobject]@{
        Timestamp=(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'); CorrelationId=$Global:CorrelationId
        Row=$rowNum; UPN=$upn; Status=$statusVal; Actions=($done -join '|'); Message=$msg
    })
    if     ($statusVal -eq 'OK')      { $ok++ }
    elseif ($statusVal -eq 'PARTIAL') { $ok++; $partial++ }
    else                              { $fail++; $failed.Add([pscustomobject]@{ $UpnCol = $upn }) }

    if (($rowNum % 25) -eq 0) {
        Write-Host "[HEARTBEAT] $($j+1)/$($plan.Count) - OK $ok / FAIL $fail" -ForegroundColor DarkGray
    }
    if (($j+1) % 100 -eq 0) {
        if (-not $headerWritten) { $results | Export-Csv -LiteralPath $Global:LogFile -NoTypeInformation -Encoding UTF8; $headerWritten=$true }
        else                     { $results | Export-Csv -LiteralPath $Global:LogFile -NoTypeInformation -Encoding UTF8 -Append }
        $results.Clear()
    }
}

# Final flush
if ($results.Count -gt 0) {
    if (-not $headerWritten) { $results | Export-Csv -LiteralPath $Global:LogFile -NoTypeInformation -Encoding UTF8 }
    else                     { $results | Export-Csv -LiteralPath $Global:LogFile -NoTypeInformation -Encoding UTF8 -Append }
}
if ($failed.Count -gt 0) { $failed | Export-Csv -LiteralPath $Global:FailedFile -NoTypeInformation -Encoding UTF8 }
if ($pwOut.Count -gt 0)  {
    $pwOut | Export-Csv -LiteralPath $Global:PwFile -NoTypeInformation -Encoding UTF8
    Write-Log "Reset passwords written to $Global:PwFile - DELETE this file after distributing." 'SECURITY'
}

$dur = (Get-Date) - $start
Write-Host ""
Write-Host "=== EXECUTION SUMMARY ===" -ForegroundColor Cyan
Write-Host ("Correlation ID  : {0}" -f $Global:CorrelationId)
Write-Host ("Planned         : {0}" -f $plan.Count)
Write-Host ("Succeeded       : {0}" -f $ok)   -ForegroundColor Green
Write-Host ("  of which partial: {0}" -f $partial) -ForegroundColor $(if($partial){'Yellow'}else{'Gray'})
Write-Host ("Failed          : {0}" -f $fail) -ForegroundColor $(if($fail){'Red'}else{'Gray'})
Write-Host ("Skipped (Pass1) : {0}" -f $skip)
Write-Host ("Duration        : {0:hh\:mm\:ss}" -f $dur)
Write-Host ("Log file        : {0}" -f $Global:LogFile)
if ($failed.Count -gt 0) { Write-Host ("Failed CSV      : {0}" -f $Global:FailedFile) -ForegroundColor Yellow }
if ($pwOut.Count -gt 0)  { Write-Host ("Passwords CSV   : {0}  (DELETE AFTER USE)" -f $Global:PwFile) -ForegroundColor Magenta }
Write-Host "=========================" -ForegroundColor Cyan
