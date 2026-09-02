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
    start/end date typed in natural formats (e.g. "August 18, 2026", "Aug 18 2026",
    "2026-08-18", "8/18/2026"). Explicit dates are read as Manila (PHT, UTC+8)
    calendar days. The window is sliced into <=ChunkDays chunks and each chunk is
    paged with ReturnLargeSet so large tenants stay under the 50,000-record
    per-session ceiling.

    Timestamps in the two CSVs are shown in Manila time (PHT, UTC+8). The audit log
    itself stores UTC; the raw JSONL retains the original UTC values for audit.

    Outputs timestamped files:
      - TeamsMeetings_*.csv             one row per meeting (organizer, start, end, count)
      - TeamsMeetingParticipants_*.csv  one row per participant per meeting (join/leave)
      - TeamsMeetingsWide_*.csv         one row per meeting: Host UPN | Display Name |
                                        Start | End | MeetingId | Count | Participant 1..N (UPNs)
      - TeamsMeetingRaw_*.jsonl         raw AuditData for every kept record (forensics)
      - TeamsMeetingTranscript_*.log    full console transcript of the run

.AUTHOR         Generated with Claude for Rakso CT Education IT.
.VERSION        1.3.1
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
                v1.1.0 - Explicit dates now accept natural formats (e.g.
                         "August 18, 2026") and are interpreted as Manila (PHT,
                         UTC+8) calendar days. CSV timestamp columns renamed *PHT
                         and converted to Manila time; raw JSONL still holds UTC.
                v1.2.0 - Hardened Search-UnifiedAuditLog against the "Failed to
                         process request via Sync Search mode ... BadRequest"
                         rejection: dates passed as Unspecified-kind UTC wall-clock
                         (no 'Z'/offset), warnings captured instead of swallowed,
                         automatic fallback to a non-paged search, and a Phase 2
                         audit-search self-test so a silent 0 can't recur.
                v1.2.1 - Added a full-session transcript (TeamsMeetingTranscript_*.log)
                         so a crash or a window that closes on file-run is captured;
                         per-record try/catch so one malformed audit record can't kill
                         the run; and per-pass progress logging (Pass A / Pass B counts)
                         so long chunks visibly advance.
                v1.3.0 - Added TeamsMeetingsWide_*.csv: one row per meeting with the
                         host UPN + host Display Name, then each participant's UPN in
                         its own column (Participant 1..N). Host display name resolved
                         from participant audit data, with a Get-Recipient fallback.
                v1.3.1 - Corrected extraction against real audit schema: Host UPN now
                         from MeetingDetail.UserId (clean UPN, not the Organizer object);
                         attendee roster read from MeetingDetail.Members[] (UPN +
                         DisplayName + Role) and merged with MeetingParticipantDetail
                         (join/leave); per-UPN dedup/merge; added Subject column; added
                         Get-Str (scalar-only) so identity objects never stringify.
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
# Returns whatever the property holds (string, number, array, object).
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

# Like Get-Prop but only accepts a SCALAR string/number - skips nested objects/arrays
# so an identity object (e.g. Organizer = @{OrganizationId=..;UserObjectId=..}) never
# gets stringified into "@{...}". Use this for UPNs, ids, dates, roles.
function Get-Str {
    param($Object,[string[]]$Names,[string]$Default='')
    $v = Get-Prop $Object $Names $null
    if ($null -eq $v) { return $Default }
    if ($v -is [string])   { return $v }
    if ($v -is [ValueType]){ return [string]$v }
    return $Default
}

# Map a Teams member/attendee Role code to a label.
function Convert-Role {
    param([string]$R)
    switch ($R) {
        '2'     { 'Organizer' }
        '1'     { 'Presenter/Attendee' }
        '3'     { 'Guest' }
        ''      { '' }
        default { "Role $R" }
    }
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

# Manila time = UTC+8, no DST. Convert a UTC datetime to PHT for display.
function ConvertTo-Pht { param($Utc) if ($null -eq $Utc) { return $null } return ([datetime]$Utc).AddHours(8) }
function Format-Pht   { param($Utc) if ($null -eq $Utc) { return '' } return (([datetime]$Utc).AddHours(8)).ToString('yyyy-MM-dd HH:mm:ss') }

# Parse a human-typed date. Accepts "August 18, 2026", "Aug 18 2026", "18 August 2026",
# "2026-08-18", "8/18/2026", etc. Returns a date (midnight, Kind=Unspecified) or $null.
function Parse-UserDate {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    $t = (Sanitize-Path $Text).Trim().TrimEnd('.')
    $formats = @(
        'MMMM d, yyyy','MMMM d yyyy','MMMM d,yyyy','MMM d, yyyy','MMM d yyyy',
        'd MMMM yyyy','d MMM yyyy','yyyy-MM-dd','yyyy/MM/dd','M/d/yyyy','MM/dd/yyyy','d-MMM-yyyy'
    )
    $ci = [cultureinfo]::GetCultureInfo('en-US')
    $dt = [datetime]::MinValue
    foreach ($f in $formats) {
        if ([datetime]::TryParseExact($t,$f,$ci,[System.Globalization.DateTimeStyles]::None,[ref]$dt)) { return $dt.Date }
    }
    if ([datetime]::TryParse($t,$ci,[System.Globalization.DateTimeStyles]::None,[ref]$dt)) { return $dt.Date }
    return $null
}

# One raw Search-UnifiedAuditLog call. Real errors throw (so Invoke-WithRetry can
# retry transient ones); warnings are captured for inspection.
function Invoke-AuditSearch {
    param($StartUnspec,$EndUnspec,[string]$Operation,[string[]]$UserIds,[string]$SessionId,[string]$SessionCommand,[int]$ResultSize)
    $wv = $null
    $p = @{ StartDate=$StartUnspec; EndDate=$EndUnspec; Operations=$Operation; ResultSize=$ResultSize; ErrorAction='Stop' }
    if ($UserIds -and $UserIds.Count -gt 0) { $p['UserIds'] = $UserIds }
    if (-not [string]::IsNullOrEmpty($SessionId)) { $p['SessionId']=$SessionId; $p['SessionCommand']=$SessionCommand }
    $recs = Search-UnifiedAuditLog @p -WarningVariable wv -WarningAction SilentlyContinue
    return @{ Records=@($recs); Warnings=@($wv) }
}

# True when a warning indicates the backend rejected the request (sync-mode BadRequest etc.).
function Test-AuditWarningFail {
    param($Warnings)
    if ($Warnings) { foreach ($w in $Warnings) { if ("$w" -match 'Sync Search mode|BadRequest|Failed to process request|Internal Server Error') { return $true } } }
    return $false
}

# Page one operation over one window; dedup by Identity. Passes UTC wall-clock as
# Unspecified kind (no 'Z'/offset for the backend to reject) and, if the paged
# ReturnLargeSet sync path is refused, falls back to a single bounded search and
# surfaces the real error instead of returning a silent 0.
function Search-AuditPaged {
    param(
        [Parameter(Mandatory)][datetime]$StartUtc,
        [Parameter(Mandatory)][datetime]$EndUtc,
        [Parameter(Mandatory)][string]$Operation,
        [string[]]$UserIds,
        [int]$ResultSize = 5000,
        [int]$MaxLoops   = 60
    )
    $sdt = [datetime]::SpecifyKind($StartUtc, [System.DateTimeKind]::Unspecified)
    $edt = [datetime]::SpecifyKind($EndUtc,   [System.DateTimeKind]::Unspecified)

    $out  = [System.Collections.Generic.List[object]]::new()
    $seen = [System.Collections.Generic.HashSet[string]]::new()
    $sid  = [guid]::NewGuid().ToString()
    $loop = 0
    $syncFailed = $false

    while ($true) {
        $loop++
        if ($loop -gt $MaxLoops) { Write-Log "Hit MaxLoops ($MaxLoops) for $Operation - reduce ChunkDays to avoid the 50k ceiling." 'WARN'; break }
        try {
            $res = Invoke-WithRetry { Invoke-AuditSearch -StartUnspec $sdt -EndUnspec $edt -Operation $Operation -UserIds $UserIds -SessionId $sid -SessionCommand 'ReturnLargeSet' -ResultSize $ResultSize }
        } catch { Write-Log "Paged search error for ${Operation}: $($_.Exception.Message)" 'WARN'; $syncFailed = $true; break }
        if (Test-AuditWarningFail $res.Warnings) { Write-Log "Paged search rejected for ${Operation}: $($res.Warnings -join ' | ')" 'WARN'; $syncFailed = $true; break }
        $batch = $res.Records
        $count = @($batch).Count
        if ($count -eq 0) { break }
        foreach ($r in $batch) {
            $id = [string](Get-Prop $r @('Identity'))
            if ([string]::IsNullOrEmpty($id)) { $id = [guid]::NewGuid().ToString() }
            if ($seen.Add($id)) { $out.Add($r) }
        }
        if ($count -lt $ResultSize) { break }
    }

    if ($syncFailed -and $out.Count -eq 0) {
        Write-Log "Falling back to non-paged search for $Operation (max $ResultSize rows this window)." 'WARN'
        try {
            $res2 = Invoke-WithRetry { Invoke-AuditSearch -StartUnspec $sdt -EndUnspec $edt -Operation $Operation -UserIds $UserIds -SessionId '' -SessionCommand '' -ResultSize $ResultSize }
            if (Test-AuditWarningFail $res2.Warnings) {
                Write-Log "Non-paged search ALSO rejected for ${Operation}: $($res2.Warnings -join ' | ')" 'ERROR'
            } else {
                foreach ($r in $res2.Records) {
                    $id = [string](Get-Prop $r @('Identity'))
                    if ([string]::IsNullOrEmpty($id)) { $id = [guid]::NewGuid().ToString() }
                    if ($seen.Add($id)) { $out.Add($r) }
                }
                if (@($res2.Records).Count -ge $ResultSize) { Write-Log "Fallback hit the $ResultSize cap for $Operation - narrow ChunkDays for full coverage." 'WARN' }
            }
        } catch { Write-Log "Non-paged search error for ${Operation}: $($_.Exception.Message)" 'ERROR' }
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

# --- Date range (explicit dates are entered in Manila time, PHT = UTC+8) ---
Write-Host ""
Write-Host "Date range:" -ForegroundColor Cyan
Write-Host "  Enter a number of DAYS to look back from now (e.g. 30, 90, 180)," -ForegroundColor Gray
Write-Host "  or press Enter to type explicit Start/End dates (Manila time)." -ForegroundColor Gray
$DaysInput = (Read-Host "Days to look back").Trim()

$StartUtc = $null; $EndUtc = $null
if ($DaysInput -match '^\d+$' -and [int]$DaysInput -gt 0) {
    $EndUtc   = (Get-Date).ToUniversalTime()
    $StartUtc = $EndUtc.AddDays(-[int]$DaysInput)
} else {
    Write-Host "  Accepted: 'August 18, 2026'  |  'Aug 18 2026'  |  '2026-08-18'  |  '8/18/2026'" -ForegroundColor DarkGray
    $sIn = (Read-Host "  Start date").Trim()
    $eIn = (Read-Host "  End date (inclusive)").Trim()
    $sD = Parse-UserDate $sIn
    $eD = Parse-UserDate $eIn
    if ($null -eq $sD -or $null -eq $eD) {
        Write-Host "ERROR: Could not read a date. Try 'August 18, 2026' or '2026-08-18'." -ForegroundColor Red; return
    }
    # Typed dates are Manila-local calendar days; convert the day boundaries to UTC (PHT = UTC+8).
    $StartUtc = [datetime]::SpecifyKind($sD.Date.AddHours(-8),               [System.DateTimeKind]::Utc)
    $EndUtc   = [datetime]::SpecifyKind($eD.Date.AddDays(1).AddHours(-8),    [System.DateTimeKind]::Utc)  # inclusive end (PHT day)
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
$WideCsv      = Join-Path $LogFolder "TeamsMeetingsWide_$TimeStamp.csv"
$RawJsonl     = Join-Path $LogFolder "TeamsMeetingRaw_$TimeStamp.jsonl"
$Global:JsonLogFile = Join-Path $LogFolder "TeamsMeetingRun_$TimeStamp.log.jsonl"
$TranscriptFile     = Join-Path $LogFolder "TeamsMeetingTranscript_$TimeStamp.log"

# Full-session transcript so a crash (or a window that closes on file-run) is never
# lost - the error lands in this file even if the console disappears.
try { Start-Transcript -LiteralPath $TranscriptFile -Force -ErrorAction Stop | Out-Null; $Global:TranscriptOn = $true }
catch { $Global:TranscriptOn = $false; Write-Host "[WARN] Could not start transcript: $($_.Exception.Message)" -ForegroundColor DarkYellow }

Write-Host ""
Write-Host "Configured:" -ForegroundColor Green
Write-Host "  Correlation ID : $Global:CorrelationId" -ForegroundColor Green
Write-Host "  Transcript     : $TranscriptFile" -ForegroundColor Green
Write-Host "  Users CSV      : $CsvPath" -ForegroundColor Green
Write-Host ("  Window (PHT)   : {0}  ->  {1}  ({2} days)" -f (Format-Pht $StartUtc), (Format-Pht $EndUtc), $totalDays) -ForegroundColor Green
Write-Host ("  Window (UTC)   : {0}  ->  {1}" -f (Format-Utc $StartUtc), (Format-Utc $EndUtc)) -ForegroundColor DarkGray
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

# Audit-search health probe: proves the search backend answers before Phase 5,
# so a silent 0 is never a mystery again.
Write-Host ""
Write-Host "Running audit-search self-test..." -ForegroundColor DarkGray
try {
    $tSt = [datetime]::SpecifyKind((Get-Date).ToUniversalTime().AddDays(-2), [System.DateTimeKind]::Unspecified)
    $tEn = [datetime]::SpecifyKind((Get-Date).ToUniversalTime(),             [System.DateTimeKind]::Unspecified)
    $tw = $null
    $probe = Search-UnifiedAuditLog -StartDate $tSt -EndDate $tEn -ResultSize 1 -WarningVariable tw -WarningAction SilentlyContinue -ErrorAction Stop
    if ($tw -and ("$tw" -match 'Sync Search mode|BadRequest|Failed to process request|Internal Server Error')) {
        Write-Host "[WARN] Audit search self-test hit a backend rejection:" -ForegroundColor Yellow
        Write-Host "       $($tw -join ' | ')" -ForegroundColor DarkYellow
        Write-Host "       This is usually an intermittent Purview backend issue (BadRequest via Sync" -ForegroundColor DarkYellow
        Write-Host "       Search mode). Phase 5 will try a non-paged fallback per chunk. If results are" -ForegroundColor DarkYellow
        Write-Host "       still 0 everywhere, wait ~15-30 min and re-run; if it persists for hours, raise" -ForegroundColor DarkYellow
        Write-Host "       a Microsoft support ticket (tenant-side fix)." -ForegroundColor DarkYellow
    } else {
        Write-Host "[OK] Audit search responded (self-test returned $(@($probe).Count) sample record(s))." -ForegroundColor Green
    }
} catch {
    Write-Host "[WARN] Audit search self-test error: $($_.Exception.Message)" -ForegroundColor Yellow
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
Write-Host ("  Window (PHT) : {0} -> {1}" -f (Format-Pht $StartUtc), (Format-Pht $EndUtc)) -ForegroundColor Yellow
Write-Host ("  Window (UTC) : {0} -> {1}" -f (Format-Utc $StartUtc), (Format-Utc $EndUtc)) -ForegroundColor DarkGray
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

# Collapse a meeting's participant rows to one per UPN, merging name/type/join/leave.
function Merge-Participants {
    param($List)
    $byUpn = [ordered]@{}
    foreach ($p in @($List)) {
        $u = [string]$p.Participant
        if ([string]::IsNullOrWhiteSpace($u)) { continue }
        $k = $u.ToLower()
        if (-not $byUpn.Contains($k)) {
            $byUpn[$k] = [PSCustomObject]@{ Participant=$u; ParticipantDisplayName=[string]$p.ParticipantDisplayName; ParticipantType=[string]$p.ParticipantType; JoinUtc=$p.JoinUtc; LeaveUtc=$p.LeaveUtc }
        } else {
            $e = $byUpn[$k]
            if ([string]::IsNullOrWhiteSpace($e.ParticipantDisplayName) -and -not [string]::IsNullOrWhiteSpace([string]$p.ParticipantDisplayName)) { $e.ParticipantDisplayName = [string]$p.ParticipantDisplayName }
            if ([string]::IsNullOrWhiteSpace($e.ParticipantType) -and -not [string]::IsNullOrWhiteSpace([string]$p.ParticipantType)) { $e.ParticipantType = [string]$p.ParticipantType }
            if ($null -ne $p.JoinUtc  -and ($null -eq $e.JoinUtc  -or $p.JoinUtc  -lt $e.JoinUtc))  { $e.JoinUtc  = $p.JoinUtc }
            if ($null -ne $p.LeaveUtc -and ($null -eq $e.LeaveUtc -or $p.LeaveUtc -gt $e.LeaveUtc)) { $e.LeaveUtc = $p.LeaveUtc }
        }
    }
    return @($byUpn.Values)
}

# Write raw AuditData line for forensics.
function Write-Raw { param($Record)
    try { $ad = Get-Prop $Record @('AuditData'); if ($ad) { [string]$ad | Add-Content -LiteralPath $RawJsonl -Encoding UTF8 } } catch { }
}

# Parse a MeetingDetail record into the meeting map (and its Members roster).
function Add-MeetingDetail {
    param($Record)
    $ad = Get-Prop $Record @('AuditData')
    if (-not $ad) { return }
    try { $j = $ad | ConvertFrom-Json -ErrorAction Stop } catch { return }
    $mid = Get-Str $j @('Id','MeetingId','MeetingDetailId','ThreadId')
    if ([string]::IsNullOrEmpty($mid)) { return }

    $recordUser = Get-Str $Record @('UserIds')
    $organizer  = Get-Str $j @('UserId')                  # clean organizer UPN (NOT the Organizer object)
    if ([string]::IsNullOrEmpty($organizer)) { $organizer = $recordUser }

    $startUtc = ConvertTo-Utc (Get-Str $j @('StartTime','MeetingStartTime'))
    $endUtc   = ConvertTo-Utc (Get-Str $j @('EndTime','MeetingEndTime'))

    # Roster: MeetingDetail.Members[] carries the attendee list WHEN present
    # (UPN, DisplayName, Role). Role 2 = organizer. This is the richest source.
    $members = Get-Prop $j @('Members')
    if ($members) {
        foreach ($mem in @($members)) {
            $mUpn = Get-Str $mem @('UPN','Upn','UserPrincipalName')
            if ([string]::IsNullOrEmpty($mUpn)) { continue }
            if ([string]::IsNullOrEmpty($organizer) -and (Get-Str $mem @('Role')) -eq '2') { $organizer = $mUpn }
            Add-Participant -MeetingId $mid -Row ([PSCustomObject]@{
                Participant            = $mUpn
                ParticipantDisplayName = (Get-Str $mem @('DisplayName','Name'))
                ParticipantType        = (Convert-Role (Get-Str $mem @('Role')))
                JoinUtc                = $null
                LeaveUtc               = $null
            })
        }
    }

    $MeetingById[$mid] = [PSCustomObject]@{
        MeetingId   = $mid
        Organizer   = $organizer
        MeetingType = (Get-Str $j @('CommunicationType','MeetingType','CommunicationSubType'))
        Subject     = (Get-Str $j @('ItemName'))
        StartUtc    = $startUtc
        EndUtc      = $endUtc
        JoinUrl     = (Get-Str $j @('MeetingURL','JoinUrl'))
        Source      = 'Organized'
    }
    Write-Raw $Record
}

# Parse a MeetingParticipantDetail record: the record's own participant (with real
# join/leave times) plus any attendee identities embedded in it. Deduped later by UPN.
function Add-ParticipantDetail {
    param($Record,[bool]$RestrictToKnownMeetings)
    $ad = Get-Prop $Record @('AuditData')
    if (-not $ad) { return }
    try { $j = $ad | ConvertFrom-Json -ErrorAction Stop } catch { return }
    $mid = Get-Str $j @('MeetingDetailId','Id','MeetingId','ThreadId')
    if ([string]::IsNullOrEmpty($mid)) { return }
    if ($RestrictToKnownMeetings -and -not $MeetingById.ContainsKey($mid)) { return }

    $joinUtc  = ConvertTo-Utc (Get-Str $j @('JoinTime'))
    $leaveUtc = ConvertTo-Utc (Get-Str $j @('LeaveTime'))

    # The record's own participant = a clean UPN with real join/leave times.
    $selfUpn = Get-Str $j @('UserId','ParticipantUPN')
    if ([string]::IsNullOrEmpty($selfUpn)) { $selfUpn = Get-Str $Record @('UserIds') }
    if (-not [string]::IsNullOrEmpty($selfUpn)) {
        Add-Participant -MeetingId $mid -Row ([PSCustomObject]@{
            Participant            = $selfUpn
            ParticipantDisplayName = ''
            ParticipantType        = ''
            JoinUtc                = $joinUtc
            LeaveUtc               = $leaveUtc
        })
    }

    # Attendee identities embedded in the record (UPN + DisplayName), when present.
    $attendees = Get-Prop $j @('Attendees')
    if ($attendees) {
        foreach ($a in @($attendees)) {
            $aUpn = Get-Str $a @('UPN','Upn','UserPrincipalName')
            if ([string]::IsNullOrEmpty($aUpn)) { continue }
            Add-Participant -MeetingId $mid -Row ([PSCustomObject]@{
                Participant            = $aUpn
                ParticipantDisplayName = (Get-Str $a @('DisplayName','Name'))
                ParticipantType        = (Convert-Role (Get-Str $a @('Role')))
                JoinUtc                = $joinUtc
                LeaveUtc               = $leaveUtc
            })
        }
    }
    Write-Raw $Record
}

$UpnArray = $TargetUpns.ToArray()
$chunkIdx = 0
foreach ($ch in $Chunks) {
    $chunkIdx++
    Write-Log ("Chunk $chunkIdx/$($Chunks.Count): {0} -> {1}" -f (Format-Utc $ch.Start), (Format-Utc $ch.End)) 'INFO'

    # Pass A - MeetingDetail (always filtered to the target users as organizers).
    Write-Log "  Pass A: searching MeetingDetail (organized by target users)..." 'INFO'
    $mdRecords = Search-AuditPaged -StartUtc $ch.Start -EndUtc $ch.End -Operation 'MeetingDetail' -UserIds $UpnArray
    $Counters.MeetingRecords += @($mdRecords).Count
    Write-Log "  Pass A: $(@($mdRecords).Count) MeetingDetail record(s)." 'INFO'
    foreach ($rec in $mdRecords) { try { Add-MeetingDetail -Record $rec } catch { Write-Log "  Skipped a MeetingDetail record: $($_.Exception.Message)" 'WARN' } }

    # Pass B - MeetingParticipantDetail.
    if ($ScopeMode -eq 1) {
        # All participants; keep only those tied to a target user's meeting (known so far).
        Write-Log "  Pass B: searching MeetingParticipantDetail (all participants)..." 'INFO'
        $pdRecords = Search-AuditPaged -StartUtc $ch.Start -EndUtc $ch.End -Operation 'MeetingParticipantDetail'
        $Counters.ParticipantRecords += @($pdRecords).Count
        Write-Log "  Pass B: $(@($pdRecords).Count) participant record(s)." 'INFO'
        foreach ($rec in $pdRecords) { try { Add-ParticipantDetail -Record $rec -RestrictToKnownMeetings $true } catch { Write-Log "  Skipped a participant record: $($_.Exception.Message)" 'WARN' } }
    } else {
        # Only the target users' own participation.
        Write-Log "  Pass B: searching MeetingParticipantDetail (target users only)..." 'INFO'
        $pdRecords = Search-AuditPaged -StartUtc $ch.Start -EndUtc $ch.End -Operation 'MeetingParticipantDetail' -UserIds $UpnArray
        $Counters.ParticipantRecords += @($pdRecords).Count
        Write-Log "  Pass B: $(@($pdRecords).Count) participant record(s)." 'INFO'
        foreach ($rec in $pdRecords) {
          try {
            # In mode 2, attended meetings may have no MeetingDetail row - register a stub.
            $adj = Get-Prop $rec @('AuditData')
            if ($adj) {
                try { $jj = $adj | ConvertFrom-Json -ErrorAction Stop } catch { $jj = $null }
                if ($jj) {
                    $mid2 = [string](Get-Prop $jj @('MeetingDetailId','Id','MeetingId','ThreadId'))
                    if (-not [string]::IsNullOrEmpty($mid2) -and -not $MeetingById.ContainsKey($mid2)) {
                        $MeetingById[$mid2] = [PSCustomObject]@{
                            MeetingId=$mid2; Organizer=(Get-Str $jj @('OrganizerUPN','OrganizerId')); MeetingType='';
                            Subject=(Get-Str $jj @('ItemName'));
                            StartUtc=(ConvertTo-Utc (Get-Str $jj @('StartTime'))); EndUtc=(ConvertTo-Utc (Get-Str $jj @('EndTime')));
                            JoinUrl=''; Source='Attended'
                        }
                    }
                }
            }
            Add-ParticipantDetail -Record $rec -RestrictToKnownMeetings $false
          } catch { Write-Log "  Skipped a participant record: $($_.Exception.Message)" 'WARN' }
        }
    }

    Write-Log ("  running totals: meetings=$($MeetingById.Count), participant records=$($Counters.ParticipantRecords)") 'INFO'
}

# --- Build output rows ---
$MeetingRows = [System.Collections.Generic.List[object]]::new()
$PartRows    = [System.Collections.Generic.List[object]]::new()

foreach ($mid in $MeetingById.Keys) {
    $m = $MeetingById[$mid]
    $raw = if ($PartsByMeeting.ContainsKey($mid)) { $PartsByMeeting[$mid] } else { @() }
    $merged = Merge-Participants $raw
    $subject = if ($m.PSObject.Properties.Name -contains 'Subject') { $m.Subject } else { '' }

    $MeetingRows.Add([PSCustomObject]@{
        CorrelationId    = $Global:CorrelationId
        MeetingId        = $m.MeetingId
        Organizer        = $m.Organizer
        Subject          = $subject
        MeetingType      = $m.MeetingType
        StartTimePHT     = (Format-Pht $m.StartUtc)
        EndTimePHT       = (Format-Pht $m.EndUtc)
        DurationMin      = (Get-DurationMin $m.StartUtc $m.EndUtc)
        ParticipantCount = @($merged).Count
        JoinUrl          = $m.JoinUrl
        Source           = $m.Source
    })

    foreach ($p in $merged) {
        $PartRows.Add([PSCustomObject]@{
            CorrelationId          = $Global:CorrelationId
            MeetingId              = $m.MeetingId
            Organizer              = $m.Organizer
            Subject                = $subject
            MeetingStartPHT        = (Format-Pht $m.StartUtc)
            MeetingEndPHT          = (Format-Pht $m.EndUtc)
            Participant            = $p.Participant
            ParticipantDisplayName = $p.ParticipantDisplayName
            ParticipantType        = $p.ParticipantType
            JoinTimePHT            = (Format-Pht $p.JoinUtc)
            LeaveTimePHT           = (Format-Pht $p.LeaveUtc)
            ParticipantDurationMin = (Get-DurationMin $p.JoinUtc $p.LeaveUtc)
        })
    }
}
$Counters.Meetings       = $MeetingRows.Count
$Counters.ParticipantRows = $PartRows.Count

# --- Build WIDE view: one row per meeting = host + participant UPNs across columns ---
# UPN -> DisplayName map harvested for free from participant audit data (Attendees carry
# both UPN and DisplayName); falls back to an Exchange Online Get-Recipient lookup (cached).
$NameMap = @{}
foreach ($k in $PartsByMeeting.Keys) {
    foreach ($p in $PartsByMeeting[$k]) {
        $pu = ([string]$p.Participant).ToLower()
        $pd = [string]$p.ParticipantDisplayName
        if (-not [string]::IsNullOrWhiteSpace($pu) -and -not [string]::IsNullOrWhiteSpace($pd) -and -not $NameMap.ContainsKey($pu)) { $NameMap[$pu] = $pd }
    }
}
$ResolveCache = @{}
function Resolve-DisplayName {
    param([string]$Upn)
    if ([string]::IsNullOrWhiteSpace($Upn)) { return '' }
    $key = $Upn.ToLower()
    if ($NameMap.ContainsKey($key))      { return $NameMap[$key] }
    if ($ResolveCache.ContainsKey($key)) { return $ResolveCache[$key] }
    $dn = ''
    try { $r = Get-Recipient -Identity $Upn -ErrorAction Stop; if ($r) { $dn = [string](Get-Prop $r @('DisplayName')) } } catch { $dn = '' }
    $ResolveCache[$key] = $dn
    return $dn
}

$wideTmp  = [System.Collections.Generic.List[object]]::new()
$maxParts = 0
foreach ($mid in $MeetingById.Keys) {
    $m = $MeetingById[$mid]
    $hostUpn = [string]$m.Organizer
    $hostKey = $hostUpn.ToLower()
    $plist = if ($PartsByMeeting.ContainsKey($mid)) { $PartsByMeeting[$mid] } else { @() }
    $distinct = [System.Collections.Generic.List[string]]::new()
    foreach ($p in $plist) {
        $pu = [string]$p.Participant
        if ([string]::IsNullOrWhiteSpace($pu)) { continue }
        if ($pu.ToLower() -eq $hostKey) { continue }         # host is already columns A/B
        if (-not ($distinct -contains $pu)) { $distinct.Add($pu) }
    }
    if ($distinct.Count -gt $maxParts) { $maxParts = $distinct.Count }
    $wideTmp.Add([PSCustomObject]@{ HostUpn=$hostUpn; HostDisp=(Resolve-DisplayName $hostUpn); MeetingId=$mid; Start=(Format-Pht $m.StartUtc); End=(Format-Pht $m.EndUtc); Parts=$distinct })
}
$WideRows = [System.Collections.Generic.List[object]]::new()
foreach ($w in $wideTmp) {
    $row = [ordered]@{
        'Host UPN'          = $w.HostUpn
        'Display Name'      = $w.HostDisp
        'Start (PHT)'       = $w.Start
        'End (PHT)'         = $w.End
        'MeetingId'         = $w.MeetingId
        'Participant Count' = $w.Parts.Count
    }
    for ($i = 0; $i -lt $maxParts; $i++) {
        $row["Participant $($i + 1)"] = if ($i -lt $w.Parts.Count) { $w.Parts[$i] } else { '' }
    }
    $WideRows.Add([PSCustomObject]$row)
}

# --- Export ---
try {
    if ($MeetingRows.Count -gt 0) { $MeetingRows | Sort-Object StartTimePHT | Export-Csv -LiteralPath $MeetingsCsv -NoTypeInformation -Encoding UTF8 }
    else { 'CorrelationId,MeetingId,Organizer,Subject,MeetingType,StartTimePHT,EndTimePHT,DurationMin,ParticipantCount,JoinUrl,Source' | Set-Content -LiteralPath $MeetingsCsv -Encoding UTF8 }
    if ($PartRows.Count -gt 0) { $PartRows | Sort-Object MeetingId, JoinTimePHT | Export-Csv -LiteralPath $PartsCsv -NoTypeInformation -Encoding UTF8 }
    else { 'CorrelationId,MeetingId,Organizer,Subject,MeetingStartPHT,MeetingEndPHT,Participant,ParticipantDisplayName,ParticipantType,JoinTimePHT,LeaveTimePHT,ParticipantDurationMin' | Set-Content -LiteralPath $PartsCsv -Encoding UTF8 }
    if ($WideRows.Count -gt 0) { $WideRows | Sort-Object 'Host UPN','Start (PHT)' | Export-Csv -LiteralPath $WideCsv -NoTypeInformation -Encoding UTF8 }
    else { 'Host UPN,Display Name,Start (PHT),End (PHT),MeetingId,Participant Count' | Set-Content -LiteralPath $WideCsv -Encoding UTF8 }
    Write-Log "Export complete." 'SUCCESS'
} catch { Write-Host "ERROR: Export failed: $($_.Exception.Message)" -ForegroundColor Red }

# --- Summary dashboard ---
$dur = (Get-Date) - $RunStart
Write-Host ""
Write-Host "=== EXECUTION SUMMARY ===" -ForegroundColor Cyan
Write-Host (" Correlation ID       : {0}" -f $Global:CorrelationId)
Write-Host (" Window (PHT)         : {0} -> {1}" -f (Format-Pht $StartUtc), (Format-Pht $EndUtc))
Write-Host (" Window (UTC)         : {0} -> {1}" -f (Format-Utc $StartUtc), (Format-Utc $EndUtc))
Write-Host (" Scope mode           : {0}" -f $ScopeMode)
Write-Host (" MeetingDetail recs   : {0}" -f $Counters.MeetingRecords)
Write-Host (" Participant recs     : {0}" -f $Counters.ParticipantRecords)
Write-Host (" Meetings (rows)      : {0}" -f $Counters.Meetings)            -ForegroundColor Green
Write-Host (" Participant rows     : {0}" -f $Counters.ParticipantRows)     -ForegroundColor Green
Write-Host (" Duration             : {0:hh\:mm\:ss}" -f $dur)
Write-Host (" Meetings CSV         : {0}" -f $MeetingsCsv)
Write-Host (" Participants CSV     : {0}" -f $PartsCsv)
Write-Host (" Wide (host+parts)    : {0}" -f $WideCsv)
Write-Host (" Raw audit (JSONL)    : {0}" -f $RawJsonl)
Write-Host " CSV times are Manila (PHT, UTC+8). Raw JSONL keeps UTC." -ForegroundColor DarkGray
Write-Host "=========================" -ForegroundColor Cyan
if ($Counters.Meetings -eq 0) {
    Write-Host "[NOTE] 0 meetings. Common causes: window older than retention; audit not yet" -ForegroundColor Yellow
    Write-Host "       ingested (can lag ~30-60 min); users organized no Teams meetings; or the" -ForegroundColor Yellow
    Write-Host "       account lacks the audit-reader role." -ForegroundColor Yellow
}

# --- Disconnect ---
try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null; Write-Host "Disconnected from Exchange Online." -ForegroundColor DarkGray } catch { }

# --- Stop transcript ---
if ((Get-Variable -Name TranscriptOn -Scope Global -ErrorAction SilentlyContinue) -and $Global:TranscriptOn) {
    try { Stop-Transcript | Out-Null; Write-Host "Transcript saved: $TranscriptFile" -ForegroundColor DarkGray } catch { }
}
