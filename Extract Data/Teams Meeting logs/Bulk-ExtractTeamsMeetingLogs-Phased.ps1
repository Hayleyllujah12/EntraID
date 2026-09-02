<#
============================================================
Bulk Extract Teams Meeting Logs (Unified Audit Log) - Phased Execution
============================================================
.SYNOPSIS
    Extracts Microsoft Teams meeting activity for specific users from the
    Microsoft 365 Unified Audit Log, filtered by a date range.

.DESCRIPTION
    Read-only reporting tool. For a CSV list of target users it pulls two
    Unified Audit Log operations and joins them:

      - MeetingDetail            -> meeting created, organizer, start, end, type, join URL
      - MeetingParticipantDetail -> participants, join time, leave time, role

    Two participant-scope modes:
      [1] Meetings ORGANIZED by the target users, with ALL participants of those
          meetings (two-pass: MeetingDetail filtered to the target users as
          organizers, then every MeetingParticipantDetail for those meeting IDs).
      [2] Only the target users' OWN participation (single-pass, both operations
          filtered by the target UPNs). Lighter; captures meetings they organized
          AND meetings they merely attended, but only their own join/leave rows.

    Date range is either "N days back from now" (N may exceed 90) or an explicit
    start/end date. The window is sliced into <=ChunkDays chunks and each chunk is
    paged with ReturnLargeSet so large tenants stay under the 50,000-record
    per-session ceiling. All timestamps are UTC (that is how the audit log stores
    them).

    Outputs three timestamped files:
      - TeamsMeetings_*.csv             one row per meeting
      - TeamsMeetingParticipants_*.csv  one row per participant per meeting
      - TeamsMeetingRaw_*.jsonl         raw AuditData for every kept record (forensics)

.AUTHOR         Rakso CT Education IT.
.VERSION        1.0.1
.DATE           2026-08-20
.REQUIREMENTS   PowerShell 7+; ExchangeOnlineManagement 3.x (Search-UnifiedAuditLog).
.PERMISSIONS    Exchange Online / Purview role: "View-Only Audit Logs" or "Audit Logs".
                Read-only. No mailbox, user, or meeting object is modified.
.SAFETY         Read-only. Only Search-UnifiedAuditLog (a query) is called.
.RETENTION      Audit (Standard) keeps records ~180 days; Audit (Premium)/E5 keeps
                ~1 year (extendable to 10 years). You can only query as far back as
                your tenant retains. Querying beyond retention returns nothing.
.CHANGELOG      v1.0   - Initial release.
                v1.0.1 - Fixed Sanitize-Path/Sanitize-Upn: String.Replace(char,'')
                         bound to the (char,char) overload and threw on the empty
                         replacement. Now casts the invisible char to string so it
                         is removed, not swapped. Broke Phase 1 on the path prompt.
============================================================
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ============================================================
# PHASE 0 - Helper functions (paste this block first)
# ============================================================

# Strip surrounding quotes and invisible Unicode from a typed path.
function Sanitize-Path {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
    $p = $Path.Trim().Trim('"').Trim("'")
    $invisible = @([char]0x00A0,[char]0x202A,[char]0x202B,[char]0x202C,[char]0x202D,
                   [char]0x202E,[char]0x200E,[char]0x200F,[char]0xFEFF,[char]0x200B)
    foreach ($c in $invisible) { $p = $p.Replace([string]$c,'') }
    return $p.Trim()
}

# Clean a UPN pulled from CSV: strip invisible chars/space variants, NFD-fold
# accents (ñ -> n), lowercase, validate local@domain. Returns a hashtable.
function Sanitize-Upn {
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
    foreach ($c in $invisible) { $s = $s.Replace([string]$c,'') }
    $s = $s -replace '\p{C}',''
    $s = $s.Trim().ToLowerInvariant()
    $changed = ($s -ne $raw)
    if ($s -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') { return @{ Clean=$s; Changed=$changed; Reason='Invalid UPN format' } }
    return @{ Clean=$s; Changed=$changed; Reason=$(if ($changed) { 'Sanitized' } else { 'Unchanged' }) }
}

# Import (prompt-before-install) a module. Never uses Update-Module.
function Ensure-Module {
    param([Parameter(Mandatory)][string]$Name,[string]$MinVersion)
    $existing = Get-Module -ListAvailable -Name $Name | Sort-Object Version -Descending | Select-Object -First 1
    if ($existing) {
        if ($MinVersion -and ($existing.Version -lt [version]$MinVersion)) {
            Write-Host "[WARN] $Name $($existing.Version) is older than required $MinVersion." -ForegroundColor Yellow
            $ans = Read-Host "Uninstall all versions and reinstall clean? (Y/N)"
            if ($ans -notmatch '^(y|yes)$') { return $false }
            try {
                Get-InstalledModule "$Name" -AllVersions -ErrorAction SilentlyContinue | Uninstall-Module -Force -ErrorAction SilentlyContinue
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

# Retry transient throttling / timeouts with exponential backoff.
function Invoke-WithRetry {
    param([Parameter(Mandatory)][scriptblock]$Action,[int]$MaxAttempts=5,[int]$BaseDelaySeconds=3)
    for ($i=1; $i -le $MaxAttempts; $i++) {
        try { return & $Action }
        catch {
            $msg = $_.Exception.Message
            $status = $null
            try { if ($_.Exception.Response -and $_.Exception.Response.StatusCode) { $status = [int]$_.Exception.Response.StatusCode } } catch { }
            $retryable = ($status -in 429,500,502,503,504) -or
                         ($msg -match 'throttl|TooManyRequests|timed out|timeout|temporarily|service is unavailable|connection reset|operation has timed out')
            if (-not $retryable -or $i -eq $MaxAttempts) { throw }
            $delay = [math]::Min(60, [int]($BaseDelaySeconds * [math]::Pow(2, $i-1)))
            Write-Host "    [RETRY $i/$MaxAttempts] transient error (status=$status). Sleeping ${delay}s..." -ForegroundColor DarkYellow
            Start-Sleep -Seconds $delay
        }
    }
}

# Console + JSONL run log, correlation-stamped.
function Write-Log {
    param([string]$Message,[ValidateSet('INFO','WARN','ERROR','SUCCESS','AUDIT')][string]$Level='INFO')
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $cid = if (Get-Variable -Name CorrelationId -Scope Global -ErrorAction SilentlyContinue) { $Global:CorrelationId } else { '------------' }
    $line = "$ts [$Level] [$cid] $Message"
    $color = @{ INFO='Gray'; WARN='Yellow'; ERROR='Red'; SUCCESS='Green'; AUDIT='Cyan' }[$Level]
    Write-Host $line -ForegroundColor $color
    if ((Get-Variable -Name JsonLogFile -Scope Global -ErrorAction SilentlyContinue) -and $Global:JsonLogFile) {
        try {
            $obj = [ordered]@{ ts=$ts; level=$Level; correlationId=$cid; message=$Message }
            ($obj | ConvertTo-Json -Compress) | Add-Content -LiteralPath $Global:JsonLogFile -Encoding UTF8
        } catch { }
    }
}

# StrictMode-safe: return the first present, non-empty property from a list of names.
function Get-Prop {
    param($Object,[string[]]$Names,$Default=$null)
    if ($null -eq $Object) { return $Default }
    $propNames = $Object.PSObject.Properties.Name
    foreach ($n in $Names) {
        if ($propNames -contains $n) {
            $v = $Object.$n
            if ($null -ne $v -and "$v" -ne '') { return $v }
        }
    }
    return $Default
}

# Parse any ISO/date string to a UTC [datetime]; $null on failure.
function ConvertTo-Utc {
    param($Value)
    if ($null -eq $Value -or "$Value" -eq '') { return $null }
    try {
        $styles = [System.Globalization.DateTimeStyles]::AssumeUniversal -bor [System.Globalization.DateTimeStyles]::AdjustToUniversal
        return ([datetimeoffset]::Parse([string]$Value,[cultureinfo]::InvariantCulture,$styles)).UtcDateTime
    } catch { return $null }
}

function Format-Utc { param($Dt) if ($null -eq $Dt) { return '' } return ([datetime]$Dt).ToString('yyyy-MM-dd HH:mm:ss') }

function Get-DurationMin {
    param($StartDt,$EndDt)
    if ($null -eq $StartDt -or $null -eq $EndDt) { return '' }
    try { return [math]::Round((([datetime]$EndDt) - ([datetime]$StartDt)).TotalMinutes, 1) } catch { return '' }
}

# Page one operation over one window with ReturnLargeSet; dedup by Identity.
function Search-AuditPaged {
    param(
        [Parameter(Mandatory)][datetime]$StartUtc,
        [Parameter(Mandatory)][datetime]$EndUtc,
        [Parameter(Mandatory)][string]$Operation,
        [string[]]$UserIds,
        [int]$ResultSize = 5000,
        [int]$MaxLoops   = 60
    )
    $sid  = [guid]::NewGuid().ToString()
    $out  = [System.Collections.Generic.List[object]]::new()
    $seen = [System.Collections.Generic.HashSet[string]]::new()
    $loop = 0
    while ($true) {
        $loop++
        if ($loop -gt $MaxLoops) {
            Write-Log "Hit MaxLoops ($MaxLoops) for $Operation in this chunk - reduce ChunkDays to avoid the 50k ceiling." 'WARN'
            break
        }
        $params = @{
            StartDate      = $StartUtc
            EndDate        = $EndUtc
            Operations     = $Operation
            SessionId      = $sid
            SessionCommand = 'ReturnLargeSet'
            ResultSize     = $ResultSize
            ErrorAction    = 'Stop'
        }
        if ($UserIds -and $UserIds.Count -gt 0) { $params['UserIds'] = $UserIds }
        $batch = Invoke-WithRetry { Search-UnifiedAuditLog @params }
        $count = @($batch).Count
        if ($count -eq 0) { break }
        foreach ($r in $batch) {
            $id = [string](Get-Prop $r @('Identity','ResultIndex'))
            if ([string]::IsNullOrEmpty($id)) { $id = [guid]::NewGuid().ToString() }
            if ($seen.Add($id)) { $out.Add($r) }
        }
        if ($count -lt $ResultSize) { break }
    }
    return ,$out
}


# ============================================================
# PHASE 1 - Configure Paths, Users, Date Range
# ============================================================
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " PHASE 1: Configure Paths, Users, Date Range" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

# Resolve the script's own folder (works when run as a file OR pasted block-by-block).
$ScriptDir = ''
try { if ($PSScriptRoot) { $ScriptDir = $PSScriptRoot } } catch { }
if (-not $ScriptDir) { try { if ($MyInvocation.MyCommand.Path) { $ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path } } catch { } }
if (-not $ScriptDir) { $ScriptDir = (Get-Location).Path }

$DefaultCsvPath   = Join-Path $ScriptDir 'Bulk-ExtractTeamsMeetingLogs-Users.csv'
$DefaultLogFolder = Join-Path $ScriptDir 'Teams_Meeting_Logs'

# --- Input CSV (target users) ---
Write-Host ""
Write-Host "Input CSV must contain a 'upn' (or 'UserPrincipalName') column." -ForegroundColor Gray
Write-Host "Default CSV path:" -ForegroundColor Gray
Write-Host "  $DefaultCsvPath" -ForegroundColor DarkGray
$InputCsv = Read-Host "Enter CSV file path (or press Enter for default)"
$CsvPath  = if ([string]::IsNullOrWhiteSpace($InputCsv)) { $DefaultCsvPath } else { Sanitize-Path $InputCsv }

if (-not (Test-Path -LiteralPath $CsvPath)) {
    Write-Host "ERROR: CSV file not found at: $CsvPath" -ForegroundColor Red
    $codes = ([char[]]$CsvPath | Select-Object -First 5 | ForEach-Object { [int]$_ }) -join ','
    Write-Host "  First 5 char codes: $codes  (a normal 'C:\' starts with 67,58,92)" -ForegroundColor DarkGray
    return
}

# --- Output folder ---
Write-Host ""
Write-Host "Default log folder:" -ForegroundColor Gray
Write-Host "  $DefaultLogFolder" -ForegroundColor DarkGray
$InputLogFolder = Read-Host "Enter output folder (or press Enter for default)"
$LogFolder = if ([string]::IsNullOrWhiteSpace($InputLogFolder)) { $DefaultLogFolder } else { Sanitize-Path $InputLogFolder }
if (-not (Test-Path -LiteralPath $LogFolder)) {
    try { New-Item -Path $LogFolder -ItemType Directory -Force | Out-Null; Write-Host "Created log folder: $LogFolder" -ForegroundColor DarkGray }
    catch { Write-Host "ERROR: Could not create log folder '$LogFolder': $($_.Exception.Message)" -ForegroundColor Red; return }
}

# --- Date range ---
Write-Host ""
Write-Host "Date range (audit log stores times in UTC):" -ForegroundColor Cyan
Write-Host "  Enter a number of DAYS to look back from now (e.g. 30, 90, 180)," -ForegroundColor Gray
Write-Host "  or press Enter to type an explicit start/end date." -ForegroundColor Gray
$DaysInput = (Read-Host "Days to look back").Trim()

$StartUtc = $null; $EndUtc = $null
if ($DaysInput -match '^\d+$' -and [int]$DaysInput -gt 0) {
    $EndUtc   = (Get-Date).ToUniversalTime()
    $StartUtc = $EndUtc.AddDays(-[int]$DaysInput)
} else {
    $sIn = (Read-Host "  Start date (yyyy-MM-dd)").Trim()
    $eIn = (Read-Host "  End date   (yyyy-MM-dd, inclusive)").Trim()
    try {
        $sD = [datetime]::ParseExact($sIn,'yyyy-MM-dd',[cultureinfo]::InvariantCulture)
        $eD = [datetime]::ParseExact($eIn,'yyyy-MM-dd',[cultureinfo]::InvariantCulture)
    } catch { Write-Host "ERROR: Dates must be yyyy-MM-dd." -ForegroundColor Red; return }
    $StartUtc = [datetime]::SpecifyKind($sD.Date, [System.DateTimeKind]::Utc)
    $EndUtc   = [datetime]::SpecifyKind($eD.Date.AddDays(1), [System.DateTimeKind]::Utc)   # inclusive end day
}
$nowUtc = (Get-Date).ToUniversalTime()
if ($EndUtc -gt $nowUtc) { $EndUtc = $nowUtc }
if ($StartUtc -ge $EndUtc) { Write-Host "ERROR: Start must be before End." -ForegroundColor Red; return }

$totalDays = [math]::Round(($EndUtc - $StartUtc).TotalDays, 1)
if ($totalDays -gt 180) {
    Write-Host "[WARN] Window is $totalDays days. Standard audit retention is ~180 days; older records" -ForegroundColor Yellow
    Write-Host "       return only if your tenant has E5 / Audit (Premium) extended retention." -ForegroundColor Yellow
}

# --- Chunk size ---
Write-Host ""
Write-Host "Window is sliced into chunks to stay under the 50,000-record per-search ceiling." -ForegroundColor Gray
$ChunkInput = (Read-Host "Chunk size in days (press Enter for 1)").Trim()
$ChunkDays  = if ($ChunkInput -match '^\d+$' -and [int]$ChunkInput -gt 0) { [int]$ChunkInput } else { 1 }

# --- Participant scope mode ---
Write-Host ""
Write-Host "Participant scope:" -ForegroundColor Cyan
Write-Host "  [1] Meetings ORGANIZED by the target users + ALL their participants (recommended, heavier)" -ForegroundColor White
Write-Host "  [2] Only the target users' OWN participation (lighter)" -ForegroundColor White
$ModeInput = (Read-Host "Enter 1 or 2 (press Enter for 1)").Trim()
$ScopeMode = if ($ModeInput -eq '2') { 2 } else { 1 }

# --- Correlation ID + output file paths ---
$Global:CorrelationId = [guid]::NewGuid().ToString('N').Substring(0,12)
$TimeStamp    = Get-Date -Format 'yyyyMMdd_HHmmss'
$MeetingsCsv  = Join-Path $LogFolder "TeamsMeetings_$TimeStamp.csv"
$PartsCsv     = Join-Path $LogFolder "TeamsMeetingParticipants_$TimeStamp.csv"
$RawJsonl     = Join-Path $LogFolder "TeamsMeetingRaw_$TimeStamp.jsonl"
$Global:JsonLogFile = Join-Path $LogFolder "TeamsMeetingRun_$TimeStamp.log.jsonl"

Write-Host ""
Write-Host "Configured:" -ForegroundColor Green
Write-Host "  Correlation ID : $Global:CorrelationId" -ForegroundColor Green
Write-Host "  Users CSV      : $CsvPath" -ForegroundColor Green
Write-Host ("  Window (UTC)   : {0}  ->  {1}  ({2} days)" -f (Format-Utc $StartUtc), (Format-Utc $EndUtc), $totalDays) -ForegroundColor Green
Write-Host "  Chunk size     : $ChunkDays day(s)" -ForegroundColor Green
Write-Host "  Scope mode     : $ScopeMode" -ForegroundColor Green
Write-Host "  Meetings out   : $MeetingsCsv" -ForegroundColor Green
Write-Host "  Participants   : $PartsCsv" -ForegroundColor Green


# ============================================================
# PHASE 2 - Connect to Exchange Online (Unified Audit Log)
# ============================================================
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " PHASE 2: Connect to Exchange Online" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

if (-not (Ensure-Module -Name 'ExchangeOnlineManagement' -MinVersion '3.0.0')) {
    Write-Host "ERROR: ExchangeOnlineManagement is required. Install with:" -ForegroundColor Red
    Write-Host "  Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force" -ForegroundColor Yellow
    return
}

Write-Host ""
$AdminUpn = (Read-Host "Enter your admin UPN (optional, prefills sign-in - press Enter to skip)").Trim()

Write-Host ""
Write-Host "Sign-in mode:" -ForegroundColor Cyan
Write-Host "  [1] Interactive browser sign-in (default)" -ForegroundColor White
Write-Host "  [2] Device code (no popup - shows a code to enter in a browser)" -ForegroundColor White
Write-Host "  [3] Reuse existing Exchange Online session if already connected" -ForegroundColor White
$SignIn = (Read-Host "Enter 1, 2, or 3 (press Enter for 1)").Trim()
if ([string]::IsNullOrWhiteSpace($SignIn)) { $SignIn = '1' }

$Connected = $false
if ($SignIn -eq '3') {
    try {
        $info = Get-ConnectionInformation -ErrorAction Stop | Where-Object { $_.State -eq 'Connected' } | Select-Object -First 1
        if ($info) { $Connected = $true; Write-Host "Reusing existing session: $($info.UserPrincipalName) @ $($info.TenantId)" -ForegroundColor Green }
        else { Write-Host "No active session found - falling back to interactive." -ForegroundColor DarkYellow; $SignIn = '1' }
    } catch { Write-Host "No active session - falling back to interactive." -ForegroundColor DarkYellow; $SignIn = '1' }
}

if (-not $Connected) {
    try {
        Write-Host ""
        Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
        $connectParams = @{ ShowBanner = $false; ErrorAction = 'Stop' }
        if (-not [string]::IsNullOrWhiteSpace($AdminUpn)) { $connectParams['UserPrincipalName'] = $AdminUpn }
        if ($SignIn -eq '2') { $connectParams['Device'] = $true }
        Connect-ExchangeOnline @connectParams
        $Connected = $true
    } catch { Write-Host "ERROR: Failed to connect to Exchange Online: $($_.Exception.Message)" -ForegroundColor Red; return }
}

# Confirm the session and (best-effort) that audit ingestion is on.
try {
    $ci = Get-ConnectionInformation -ErrorAction Stop | Where-Object { $_.State -eq 'Connected' } | Select-Object -First 1
    if ($ci) {
        Write-Host ""
        Write-Host "Connected." -ForegroundColor Green
        Write-Host "  Account : $($ci.UserPrincipalName)" -ForegroundColor Green
        Write-Host "  Tenant  : $($ci.TenantId)" -ForegroundColor Green
    }
} catch { }
try {
    $cfg = Get-AdminAuditLogConfig -ErrorAction Stop
    $ingest = Get-Prop $cfg @('UnifiedAuditLogIngestionEnabled')
    if ($null -ne $ingest -and -not $ingest) {
        Write-Host "[WARN] UnifiedAuditLogIngestionEnabled = False. Audit search may return nothing." -ForegroundColor Yellow
    }
} catch {
    Write-Host "  (Could not read audit config - continuing; your role may lack Get-AdminAuditLogConfig.)" -ForegroundColor DarkGray
}


# ============================================================
# PHASE 3 - Verify Users CSV
# ============================================================
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " PHASE 3: Verify Users CSV" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

try { $Rows = Import-Csv -LiteralPath $CsvPath -ErrorAction Stop }
catch { Write-Host "ERROR: Failed to read CSV: $($_.Exception.Message)" -ForegroundColor Red; return }

if ($null -eq $Rows -or @($Rows).Count -eq 0) { Write-Host "ERROR: CSV is empty." -ForegroundColor Red; return }

$cols = $Rows[0].PSObject.Properties.Name
$UpnCol = $null
foreach ($cand in @('upn','UPN','UserPrincipalName','userprincipalname','Upn')) { if ($cols -contains $cand) { $UpnCol = $cand; break } }
if (-not $UpnCol) {
    Write-Host "ERROR: CSV must contain a 'upn' or 'UserPrincipalName' column." -ForegroundColor Red
    Write-Host "Found columns: $($cols -join ', ')" -ForegroundColor Yellow
    return
}

$TargetUpns = [System.Collections.Generic.List[string]]::new()
$changedReport = [System.Collections.Generic.List[object]]::new()
foreach ($r in $Rows) {
    $rawVal = if ($null -ne $r.$UpnCol) { [string]$r.$UpnCol } else { '' }
    $res = Sanitize-Upn -Value $rawVal
    if ($res.Reason -eq 'Invalid UPN format' -or [string]::IsNullOrWhiteSpace($res.Clean)) {
        $changedReport.Add([PSCustomObject]@{ Raw=$rawVal; Clean=$res.Clean; Note=$res.Reason }); continue
    }
    if ($res.Changed) { $changedReport.Add([PSCustomObject]@{ Raw=$rawVal; Clean=$res.Clean; Note='Sanitized' }) }
    if (-not $TargetUpns.Contains($res.Clean)) { $TargetUpns.Add($res.Clean) }
}

if ($TargetUpns.Count -eq 0) { Write-Host "ERROR: No valid UPNs after sanitization." -ForegroundColor Red; return }

Write-Host ""
Write-Host "Loaded $($TargetUpns.Count) unique valid user(s) from column '$UpnCol'." -ForegroundColor Green
$TargetUpns | Select-Object -First 20 | ForEach-Object { Write-Host "  - $_" -ForegroundColor DarkGray }
if ($TargetUpns.Count -gt 20) { Write-Host "  ... and $($TargetUpns.Count - 20) more" -ForegroundColor DarkGray }
if ($changedReport.Count -gt 0) {
    Write-Host ""
    Write-Host "Rows cleaned / skipped:" -ForegroundColor Yellow
    $changedReport | Select-Object -First 15 | Format-Table -AutoSize | Out-Host
}

# Precompute the chunk plan.
$Chunks = [System.Collections.Generic.List[object]]::new()
$cursor = $StartUtc
while ($cursor -lt $EndUtc) {
    $next = $cursor.AddDays($ChunkDays)
    if ($next -gt $EndUtc) { $next = $EndUtc }
    $Chunks.Add([PSCustomObject]@{ Start=$cursor; End=$next })
    $cursor = $next
}
Write-Host ""
Write-Host "Search plan: $($Chunks.Count) chunk(s) of up to $ChunkDays day(s) each." -ForegroundColor Green


# ============================================================
# PHASE 4 - Confirmation
# ============================================================
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " PHASE 4: Confirmation" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""
$scopeText = if ($ScopeMode -eq 1) { "meetings ORGANIZED by these users + all their participants" } else { "these users' OWN participation only" }
Write-Host "About to QUERY the Unified Audit Log (READ-ONLY):" -ForegroundColor Yellow
Write-Host "  Users        : $($TargetUpns.Count)" -ForegroundColor Yellow
Write-Host ("  Window (UTC) : {0} -> {1}" -f (Format-Utc $StartUtc), (Format-Utc $EndUtc)) -ForegroundColor Yellow
Write-Host "  Chunks       : $($Chunks.Count) x up to $ChunkDays day(s)" -ForegroundColor Yellow
Write-Host "  Scope        : $scopeText" -ForegroundColor Yellow
Write-Host "  Operations   : MeetingDetail, MeetingParticipantDetail" -ForegroundColor Yellow
Write-Host ""
Write-Host "No tenant object is modified. This only reads audit records." -ForegroundColor Green
Write-Host ""
$Confirm = Read-Host "Proceed? (Y/N)"
if ($Confirm -notmatch '^(y|yes)$') {
    Write-Host "Cancelled by user. Disconnecting..." -ForegroundColor Red
    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null } catch { }
    return
}


# ============================================================
# PHASE 5 - Extract, Join, Export
# ============================================================
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " PHASE 5: Extract Teams Meeting Activity" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

$RunStart      = Get-Date
$MeetingById   = @{}   # MeetingId -> meeting PSCustomObject
$PartsByMeeting = @{}  # MeetingId -> List[participant rows]
$Counters      = @{ MeetingRecords=0; ParticipantRecords=0; Meetings=0; ParticipantRows=0 }

# Ensure a participant bucket exists.
function Add-Participant {
    param($MeetingId,$Row)
    if (-not $PartsByMeeting.ContainsKey($MeetingId)) { $PartsByMeeting[$MeetingId] = [System.Collections.Generic.List[object]]::new() }
    $PartsByMeeting[$MeetingId].Add($Row)
}

# Write raw AuditData line for forensics.
function Write-Raw { param($Record)
    try { $ad = Get-Prop $Record @('AuditData'); if ($ad) { [string]$ad | Add-Content -LiteralPath $RawJsonl -Encoding UTF8 } } catch { }
}

# Parse a MeetingDetail record into the meeting map.
function Add-MeetingDetail {
    param($Record)
    $ad = Get-Prop $Record @('AuditData')
    if (-not $ad) { return }
    try { $j = $ad | ConvertFrom-Json -ErrorAction Stop } catch { return }
    $mid = [string](Get-Prop $j @('Id','MeetingId','MeetingDetailId','ThreadId'))
    if ([string]::IsNullOrEmpty($mid)) { return }
    $recordUser = [string](Get-Prop $Record @('UserIds'))
    $organizer  = [string](Get-Prop $j @('Organizer','UserId'))
    if ([string]::IsNullOrEmpty($organizer)) { $organizer = $recordUser }
    $startUtc = ConvertTo-Utc (Get-Prop $j @('StartTime','MeetingStartTime'))
    $endUtc   = ConvertTo-Utc (Get-Prop $j @('EndTime','MeetingEndTime'))
    $obj = [PSCustomObject]@{
        MeetingId   = $mid
        Organizer   = $organizer
        MeetingType = [string](Get-Prop $j @('MeetingType','ItemName','CommunicationSubType','Communication'))
        StartUtc    = $startUtc
        EndUtc      = $endUtc
        JoinUrl     = [string](Get-Prop $j @('MeetingURL','JoinUrl','ItemName'))
        Source      = 'Organized'
    }
    $MeetingById[$mid] = $obj
    Write-Raw $Record
}

# Parse a MeetingParticipantDetail record into participant rows (attendee expansion).
function Add-ParticipantDetail {
    param($Record,[bool]$RestrictToKnownMeetings)
    $ad = Get-Prop $Record @('AuditData')
    if (-not $ad) { return }
    try { $j = $ad | ConvertFrom-Json -ErrorAction Stop } catch { return }
    $mid = [string](Get-Prop $j @('MeetingDetailId','Id','MeetingId','ThreadId'))
    if ([string]::IsNullOrEmpty($mid)) { return }
    if ($RestrictToKnownMeetings -and -not $MeetingById.ContainsKey($mid)) { return }

    $joinUtc  = ConvertTo-Utc (Get-Prop $j @('JoinTime'))
    $leaveUtc = ConvertTo-Utc (Get-Prop $j @('LeaveTime'))
    $attendees = Get-Prop $j @('Attendees','Members','Participants')

    $emit = {
        param($upn,$disp,$ptype)
        Add-Participant -MeetingId $mid -Row ([PSCustomObject]@{
            Participant            = [string]$upn
            ParticipantDisplayName = [string]$disp
            ParticipantType        = [string]$ptype
            JoinUtc                = $joinUtc
            LeaveUtc               = $leaveUtc
        })
    }

    if ($attendees -and @($attendees).Count -gt 0) {
        foreach ($a in @($attendees)) {
            $upn = Get-Prop $a @('UPN','Upn','UserPrincipalName','Id','RecipientEmailAddress')
            $disp= Get-Prop $a @('DisplayName','Name')
            $pt  = Get-Prop $a @('RecipientType','Role','ParticipantType')
            & $emit $upn $disp $pt
        }
    } else {
        $upn = Get-Prop $j @('UserId','ParticipantUPN')
        if ([string]::IsNullOrEmpty([string](Get-Prop $Record @('UserIds')))) { } # noop
        if ([string]::IsNullOrEmpty([string]$upn)) { $upn = [string](Get-Prop $Record @('UserIds')) }
        $pt  = Get-Prop $j @('Role','RecipientType')
        & $emit $upn '' $pt
    }
    Write-Raw $Record
}

$UpnArray = $TargetUpns.ToArray()
$chunkIdx = 0
foreach ($ch in $Chunks) {
    $chunkIdx++
    Write-Log ("Chunk $chunkIdx/$($Chunks.Count): {0} -> {1}" -f (Format-Utc $ch.Start), (Format-Utc $ch.End)) 'INFO'

    # Pass A - MeetingDetail (always filtered to the target users as organizers).
    $mdRecords = Search-AuditPaged -StartUtc $ch.Start -EndUtc $ch.End -Operation 'MeetingDetail' -UserIds $UpnArray
    $Counters.MeetingRecords += @($mdRecords).Count
    foreach ($rec in $mdRecords) { Add-MeetingDetail -Record $rec }

    # Pass B - MeetingParticipantDetail.
    if ($ScopeMode -eq 1) {
        # All participants; keep only those tied to a target user's meeting (known so far).
        $pdRecords = Search-AuditPaged -StartUtc $ch.Start -EndUtc $ch.End -Operation 'MeetingParticipantDetail'
        $Counters.ParticipantRecords += @($pdRecords).Count
        foreach ($rec in $pdRecords) { Add-ParticipantDetail -Record $rec -RestrictToKnownMeetings $true }
    } else {
        # Only the target users' own participation.
        $pdRecords = Search-AuditPaged -StartUtc $ch.Start -EndUtc $ch.End -Operation 'MeetingParticipantDetail' -UserIds $UpnArray
        $Counters.ParticipantRecords += @($pdRecords).Count
        foreach ($rec in $pdRecords) {
            # In mode 2, attended meetings may have no MeetingDetail row - register a stub.
            $adj = Get-Prop $rec @('AuditData')
            if ($adj) {
                try { $jj = $adj | ConvertFrom-Json -ErrorAction Stop } catch { $jj = $null }
                if ($jj) {
                    $mid2 = [string](Get-Prop $jj @('MeetingDetailId','Id','MeetingId','ThreadId'))
                    if (-not [string]::IsNullOrEmpty($mid2) -and -not $MeetingById.ContainsKey($mid2)) {
                        $MeetingById[$mid2] = [PSCustomObject]@{
                            MeetingId=$mid2; Organizer=[string](Get-Prop $jj @('Organizer','OrganizerId')); MeetingType='';
                            StartUtc=(ConvertTo-Utc (Get-Prop $jj @('StartTime'))); EndUtc=(ConvertTo-Utc (Get-Prop $jj @('EndTime')));
                            JoinUrl=''; Source='Attended'
                        }
                    }
                }
            }
            Add-ParticipantDetail -Record $rec -RestrictToKnownMeetings $false
        }
    }

    Write-Log ("  running totals: meetings=$($MeetingById.Count), participant records=$($Counters.ParticipantRecords)") 'INFO'
}

# --- Build output rows ---
$MeetingRows = [System.Collections.Generic.List[object]]::new()
$PartRows    = [System.Collections.Generic.List[object]]::new()

foreach ($mid in $MeetingById.Keys) {
    $m = $MeetingById[$mid]
    $plist = if ($PartsByMeeting.ContainsKey($mid)) { $PartsByMeeting[$mid] } else { [System.Collections.Generic.List[object]]::new() }

    $MeetingRows.Add([PSCustomObject]@{
        CorrelationId    = $Global:CorrelationId
        MeetingId        = $m.MeetingId
        Organizer        = $m.Organizer
        MeetingType      = $m.MeetingType
        StartTimeUTC     = (Format-Utc $m.StartUtc)
        EndTimeUTC       = (Format-Utc $m.EndUtc)
        DurationMin      = (Get-DurationMin $m.StartUtc $m.EndUtc)
        ParticipantCount = @($plist).Count
        JoinUrl          = $m.JoinUrl
        Source           = $m.Source
    })

    foreach ($p in $plist) {
        $PartRows.Add([PSCustomObject]@{
            CorrelationId          = $Global:CorrelationId
            MeetingId              = $m.MeetingId
            Organizer              = $m.Organizer
            MeetingStartUTC        = (Format-Utc $m.StartUtc)
            MeetingEndUTC          = (Format-Utc $m.EndUtc)
            Participant            = $p.Participant
            ParticipantDisplayName = $p.ParticipantDisplayName
            ParticipantType        = $p.ParticipantType
            JoinTimeUTC            = (Format-Utc $p.JoinUtc)
            LeaveTimeUTC           = (Format-Utc $p.LeaveUtc)
            ParticipantDurationMin = (Get-DurationMin $p.JoinUtc $p.LeaveUtc)
        })
    }
}
$Counters.Meetings       = $MeetingRows.Count
$Counters.ParticipantRows = $PartRows.Count

# --- Export ---
try {
    if ($MeetingRows.Count -gt 0) { $MeetingRows | Sort-Object StartTimeUTC | Export-Csv -LiteralPath $MeetingsCsv -NoTypeInformation -Encoding UTF8 }
    else { 'CorrelationId,MeetingId,Organizer,MeetingType,StartTimeUTC,EndTimeUTC,DurationMin,ParticipantCount,JoinUrl,Source' | Set-Content -LiteralPath $MeetingsCsv -Encoding UTF8 }
    if ($PartRows.Count -gt 0) { $PartRows | Sort-Object MeetingId, JoinTimeUTC | Export-Csv -LiteralPath $PartsCsv -NoTypeInformation -Encoding UTF8 }
    else { 'CorrelationId,MeetingId,Organizer,MeetingStartUTC,MeetingEndUTC,Participant,ParticipantDisplayName,ParticipantType,JoinTimeUTC,LeaveTimeUTC,ParticipantDurationMin' | Set-Content -LiteralPath $PartsCsv -Encoding UTF8 }
    Write-Log "Export complete." 'SUCCESS'
} catch { Write-Host "ERROR: Export failed: $($_.Exception.Message)" -ForegroundColor Red }

# --- Summary dashboard ---
$dur = (Get-Date) - $RunStart
Write-Host ""
Write-Host "=== EXECUTION SUMMARY ===" -ForegroundColor Cyan
Write-Host (" Correlation ID       : {0}" -f $Global:CorrelationId)
Write-Host (" Window (UTC)         : {0} -> {1}" -f (Format-Utc $StartUtc), (Format-Utc $EndUtc))
Write-Host (" Scope mode           : {0}" -f $ScopeMode)
Write-Host (" MeetingDetail recs   : {0}" -f $Counters.MeetingRecords)
Write-Host (" Participant recs     : {0}" -f $Counters.ParticipantRecords)
Write-Host (" Meetings (rows)      : {0}" -f $Counters.Meetings)            -ForegroundColor Green
Write-Host (" Participant rows     : {0}" -f $Counters.ParticipantRows)     -ForegroundColor Green
Write-Host (" Duration             : {0:hh\:mm\:ss}" -f $dur)
Write-Host (" Meetings CSV         : {0}" -f $MeetingsCsv)
Write-Host (" Participants CSV     : {0}" -f $PartsCsv)
Write-Host (" Raw audit (JSONL)    : {0}" -f $RawJsonl)
Write-Host "=========================" -ForegroundColor Cyan
if ($Counters.Meetings -eq 0) {
    Write-Host "[NOTE] 0 meetings. Common causes: window older than retention; audit not yet" -ForegroundColor Yellow
    Write-Host "       ingested (can lag ~30-60 min); users organized no Teams meetings; or the" -ForegroundColor Yellow
    Write-Host "       account lacks the audit-reader role." -ForegroundColor Yellow
}

# --- Disconnect ---
try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null; Write-Host "Disconnected from Exchange Online." -ForegroundColor DarkGray } catch { }