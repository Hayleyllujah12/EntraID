<#
============================================================
Bulk Extract Teams Meeting ATTENDANCE (Graph Attendance Reports) - Phased Execution
============================================================
.SYNOPSIS
    Extracts full Teams meeting attendance (name, email, role, join/leave, duration)
    for meetings ORGANIZED by a list of users, over a date range, using the Microsoft
    Graph attendance-report API. This is the reliable source for the complete roster
    (including students) that the Unified Audit Log does NOT store for regular meetings.

.DESCRIPTION
    For each organizer UPN in the input CSV:
      1. Read their calendar (CalendarView) over the date range and keep Teams online meetings.
      2. Resolve each meeting's JoinWebUrl -> onlineMeeting id.
      3. Pull onlineMeeting attendanceReports -> attendanceRecords.
      4. Emit attendance rows (per attendee per meeting) + a wide one-row-per-meeting view.

    Auth is selected at runtime:
      [1] Delegated interactive  - reads ONLY the signed-in user's own meetings.
      [2] App-only (certificate) - reads all targeted organizers (needs app registration
                                   + a Teams application access policy). Recommended for bulk.
      [3] App-only (client secret) - same as [2] but with a secret instead of a cert.

    Times in the CSVs are Manila (PHT, UTC+8). Raw JSONL keeps the original UTC.

    Outputs (timestamped):
      - TeamsAttendance_*.csv        one row per attendee per meeting (name,email,join,leave,minutes,role)
      - TeamsAttendanceWide_*.csv    one row per meeting: Host UPN | Display Name | Subject |
                                     Start | End | MeetingId | Count | Participant 1..N (emails)
      - TeamsAttendanceRaw_*.jsonl   raw attendance records (forensics)
      - TeamsAttendanceTranscript_*.log

.AUTHOR         Generated with Claude for Rakso CT Education IT.
.VERSION        1.3.0
.DATE           2026-08-20
.REQUIREMENTS   PowerShell 7+; Microsoft.Graph.Authentication, .Calendar, .CloudCommunications.
.PERMISSIONS    Delegated: User.Read, Calendars.Read, OnlineMeetings.Read,
                           OnlineMeetingArtifact.Read.All (self only; no app registration).
                App-only : Calendars.Read, OnlineMeetings.Read.All, OnlineMeetingArtifact.Read.All
                           (all read-only). The Teams onlineMeetings API needs the organizer's
                           OBJECT ID; to let the script translate UPN -> id add User.ReadBasic.All,
                           OR skip it entirely by putting each organizer's OBJECT ID in the CSV.
                           Least privilege = scope calendars + meetings to ONE security group of
                           the organizers: Exchange New-ApplicationAccessPolicy (calendars) +
                           Teams Grant-CsApplicationAccessPolicy -Group (meetings/attendance).
                           Host display name comes from the calendar event (no directory read).
                           See setup notes at the end of the file.
.SAFETY         Read-only. No mailbox, meeting, or user object is modified.
.CHANGELOG      v1.0   - Initial release.
                v1.1.0 - Least-privilege: dropped User.Read.All (host display name now comes
                         from the calendar event organizer) and Microsoft.Graph.Users. Delegated
                         scopes trimmed to Calendars.Read, OnlineMeetings.Read,
                         OnlineMeetingArtifact.Read.All. Setup notes updated with group-scoped
                         Exchange + Teams application access policies.
                v1.2.0 - Fixed "The userId in request URL is not a valid GUID": the Teams
                         onlineMeetings API requires the organizer OBJECT ID, not the UPN. Added
                         Resolve-UserId (accepts an id directly, or translates a UPN via
                         Invoke-MgGraphRequest) and use the id for all meeting/attendance calls.
                         CSV may now contain object IDs (no directory read) or UPNs (needs
                         User.ReadBasic.All app-only). Delegated resolves self via /me (User.Read).
                v1.2.1 - Fixed attendanceRecords call: the SDK parameter is
                         -MeetingAttendanceReportId (named after the meetingAttendanceReport
                         resource), not -AttendanceReportId.
                v1.2.2 - Added per-meeting heartbeat ("Meeting X/N: resolving ...", plus a
                         record count per report) so Phase 5 visibly advances instead of going
                         silent between the calendar count and the next organizer.
                v1.3.0 - Fixed out-of-window dates leaking in: a recurring meeting is one
                         onlineMeeting whose attendanceReports span its whole history, so
                         finding one in-window occurrence pulled every past session. Now each
                         report is filtered by its MeetingStartDateTime to the requested
                         [start,end) window; skipped ones are counted in the summary.
============================================================
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ============================================================
# PHASE 0 - Helper functions (paste this block first)
# ============================================================

function Sanitize-Path {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
    $p = $Path.Trim().Trim('"').Trim("'")
    $invisible = @([char]0x00A0,[char]0x202A,[char]0x202B,[char]0x202C,[char]0x202D,
                   [char]0x202E,[char]0x200E,[char]0x200F,[char]0xFEFF,[char]0x200B)
    foreach ($c in $invisible) { $p = $p.Replace([string]$c,'') }
    return $p.Trim()
}

function Sanitize-Upn {
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return @{ Clean=''; Changed=$true; Reason='Empty or whitespace-only' } }
    $raw = $Value
    $nfd = $raw.Normalize([Text.NormalizationForm]::FormD)
    $sb  = [Text.StringBuilder]::new()
    foreach ($ch in $nfd.ToCharArray()) {
        if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne [Globalization.UnicodeCategory]::NonSpacingMark) { [void]$sb.Append($ch) }
    }
    $s = $sb.ToString()
    $invisible = @([char]0x00A0,[char]0x1680,[char]0x2000,[char]0x2001,[char]0x2002,[char]0x2003,
        [char]0x2004,[char]0x2005,[char]0x2006,[char]0x2007,[char]0x2008,[char]0x2009,
        [char]0x200A,[char]0x200B,[char]0x200C,[char]0x200D,[char]0x200E,[char]0x200F,
        [char]0x202A,[char]0x202B,[char]0x202C,[char]0x202D,[char]0x202E,
        [char]0x202F,[char]0x205F,[char]0x2060,[char]0x3000,[char]0xFEFF)
    foreach ($c in $invisible) { $s = $s.Replace([string]$c,'') }
    $s = $s -replace '\p{C}',''
    $s = $s.Trim().ToLowerInvariant()
    $changed = ($s -ne $raw)
    if ($s -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') { return @{ Clean=$s; Changed=$changed; Reason='Invalid UPN format' } }
    return @{ Clean=$s; Changed=$changed; Reason=$(if ($changed) { 'Sanitized' } else { 'Unchanged' }) }
}

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
    try { Install-Module $Name -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop; Import-Module $Name -ErrorAction Stop; return $true }
    catch { Write-Host "[FAIL] Install of $Name failed: $($_.Exception.Message)" -ForegroundColor Red; return $false }
}

function Invoke-WithRetry {
    param([Parameter(Mandatory)][scriptblock]$Action,[int]$MaxAttempts=5,[int]$BaseDelaySeconds=3)
    for ($i=1; $i -le $MaxAttempts; $i++) {
        try { return & $Action }
        catch {
            $msg = $_.Exception.Message
            $status = $null
            try { if ($_.Exception.Response -and $_.Exception.Response.StatusCode) { $status = [int]$_.Exception.Response.StatusCode } } catch { }
            $retryable = ($status -in 429,500,502,503,504) -or ($msg -match 'throttl|TooManyRequests|timed out|timeout|temporarily|service is unavailable|connection reset')
            if (-not $retryable -or $i -eq $MaxAttempts) { throw }
            $delay = [math]::Min(60, [int]($BaseDelaySeconds * [math]::Pow(2, $i-1)))
            Write-Host "    [RETRY $i/$MaxAttempts] transient error (status=$status). Sleeping ${delay}s..." -ForegroundColor DarkYellow
            Start-Sleep -Seconds $delay
        }
    }
}

function Write-Log {
    param([string]$Message,[ValidateSet('INFO','WARN','ERROR','SUCCESS','AUDIT')][string]$Level='INFO')
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $cid = if (Get-Variable -Name CorrelationId -Scope Global -ErrorAction SilentlyContinue) { $Global:CorrelationId } else { '------------' }
    $line = "$ts [$Level] [$cid] $Message"
    $color = @{ INFO='Gray'; WARN='Yellow'; ERROR='Red'; SUCCESS='Green'; AUDIT='Cyan' }[$Level]
    Write-Host $line -ForegroundColor $color
    if ((Get-Variable -Name JsonLogFile -Scope Global -ErrorAction SilentlyContinue) -and $Global:JsonLogFile) {
        try { $obj=[ordered]@{ts=$ts;level=$Level;correlationId=$cid;message=$Message}; ($obj|ConvertTo-Json -Compress)|Add-Content -LiteralPath $Global:JsonLogFile -Encoding UTF8 } catch { }
    }
}

function Get-Prop {
    param($Object,[string[]]$Names,$Default=$null)
    if ($null -eq $Object) { return $Default }
    $propNames = $Object.PSObject.Properties.Name
    foreach ($n in $Names) {
        if ($propNames -contains $n) { $v = $Object.$n; if ($null -ne $v -and "$v" -ne '') { return $v } }
    }
    return $Default
}

function Get-Str {
    param($Object,[string[]]$Names,[string]$Default='')
    $v = Get-Prop $Object $Names $null
    if ($null -eq $v) { return $Default }
    if ($v -is [string])    { return $v }
    if ($v -is [ValueType]) { return [string]$v }
    return $Default
}

function ConvertTo-Utc {
    param($Value)
    if ($null -eq $Value -or "$Value" -eq '') { return $null }
    if ($Value -is [datetime]) { return ([datetime]$Value).ToUniversalTime() }
    try {
        $styles = [System.Globalization.DateTimeStyles]::AssumeUniversal -bor [System.Globalization.DateTimeStyles]::AdjustToUniversal
        return ([datetimeoffset]::Parse([string]$Value,[cultureinfo]::InvariantCulture,$styles)).UtcDateTime
    } catch { return $null }
}
function Format-Pht { param($Utc) if ($null -eq $Utc) { return '' } return (([datetime]$Utc).AddHours(8)).ToString('yyyy-MM-dd HH:mm:ss') }
function Format-Utc { param($Utc) if ($null -eq $Utc) { return '' } return ([datetime]$Utc).ToString('yyyy-MM-dd HH:mm:ss') }

function Parse-UserDate {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    $t = (Sanitize-Path $Text).Trim().TrimEnd('.')
    $formats = @('MMMM d, yyyy','MMMM d yyyy','MMMM d,yyyy','MMM d, yyyy','MMM d yyyy',
                 'd MMMM yyyy','d MMM yyyy','yyyy-MM-dd','yyyy/MM/dd','M/d/yyyy','MM/dd/yyyy','d-MMM-yyyy')
    $ci = [cultureinfo]::GetCultureInfo('en-US'); $dt=[datetime]::MinValue
    foreach ($f in $formats) { if ([datetime]::TryParseExact($t,$f,$ci,[System.Globalization.DateTimeStyles]::None,[ref]$dt)) { return $dt.Date } }
    if ([datetime]::TryParse($t,$ci,[System.Globalization.DateTimeStyles]::None,[ref]$dt)) { return $dt.Date }
    return $null
}


# ============================================================
# PHASE 1 - Configure Paths, Organizers, Date Range
# ============================================================
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " PHASE 1: Configure Paths, Organizers, Date Range" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

$ScriptDir = ''
try { if ($PSScriptRoot) { $ScriptDir = $PSScriptRoot } } catch { }
if (-not $ScriptDir) { try { if ($MyInvocation.MyCommand.Path) { $ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path } } catch { } }
if (-not $ScriptDir) { $ScriptDir = (Get-Location).Path }

$DefaultCsvPath   = Join-Path $ScriptDir 'Bulk-ExtractTeamsAttendance-Organizers.csv'
$DefaultLogFolder = Join-Path $ScriptDir 'Teams_Attendance_Logs'

Write-Host ""
Write-Host "Input CSV must contain a 'upn' (or 'UserPrincipalName') column of MEETING ORGANIZERS." -ForegroundColor Gray
Write-Host "  Default: $DefaultCsvPath" -ForegroundColor DarkGray
$InputCsv = Read-Host "Enter CSV file path (or press Enter for default)"
$CsvPath  = if ([string]::IsNullOrWhiteSpace($InputCsv)) { $DefaultCsvPath } else { Sanitize-Path $InputCsv }
if (-not (Test-Path -LiteralPath $CsvPath)) { Write-Host "ERROR: CSV not found at: $CsvPath" -ForegroundColor Red; return }

Write-Host ""
Write-Host "  Default output folder: $DefaultLogFolder" -ForegroundColor DarkGray
$InputLogFolder = Read-Host "Enter output folder (or press Enter for default)"
$LogFolder = if ([string]::IsNullOrWhiteSpace($InputLogFolder)) { $DefaultLogFolder } else { Sanitize-Path $InputLogFolder }
if (-not (Test-Path -LiteralPath $LogFolder)) {
    try { New-Item -Path $LogFolder -ItemType Directory -Force | Out-Null } catch { Write-Host "ERROR: Could not create log folder: $($_.Exception.Message)" -ForegroundColor Red; return }
}

Write-Host ""
Write-Host "Date range (explicit dates are read as Manila / PHT):" -ForegroundColor Cyan
Write-Host "  Enter DAYS to look back (e.g. 30), or press Enter to type Start/End dates." -ForegroundColor Gray
$DaysInput = (Read-Host "Days to look back").Trim()
$StartUtc = $null; $EndUtc = $null
if ($DaysInput -match '^\d+$' -and [int]$DaysInput -gt 0) {
    $EndUtc = (Get-Date).ToUniversalTime(); $StartUtc = $EndUtc.AddDays(-[int]$DaysInput)
} else {
    Write-Host "  Accepted: 'August 18, 2026' | 'Aug 18 2026' | '2026-08-18' | '8/18/2026'" -ForegroundColor DarkGray
    $sD = Parse-UserDate ((Read-Host "  Start date").Trim())
    $eD = Parse-UserDate ((Read-Host "  End date (inclusive)").Trim())
    if ($null -eq $sD -or $null -eq $eD) { Write-Host "ERROR: Could not read a date." -ForegroundColor Red; return }
    $StartUtc = [datetime]::SpecifyKind($sD.Date.AddHours(-8),            [System.DateTimeKind]::Utc)
    $EndUtc   = [datetime]::SpecifyKind($eD.Date.AddDays(1).AddHours(-8), [System.DateTimeKind]::Utc)
}
$nowUtc = (Get-Date).ToUniversalTime()
if ($EndUtc -gt $nowUtc) { $EndUtc = $nowUtc }
if ($StartUtc -ge $EndUtc) { Write-Host "ERROR: Start must be before End." -ForegroundColor Red; return }

$Global:CorrelationId = [guid]::NewGuid().ToString('N').Substring(0,12)
$TimeStamp   = Get-Date -Format 'yyyyMMdd_HHmmss'
$AttCsv      = Join-Path $LogFolder "TeamsAttendance_$TimeStamp.csv"
$WideCsv     = Join-Path $LogFolder "TeamsAttendanceWide_$TimeStamp.csv"
$RawJsonl    = Join-Path $LogFolder "TeamsAttendanceRaw_$TimeStamp.jsonl"
$TranscriptFile     = Join-Path $LogFolder "TeamsAttendanceTranscript_$TimeStamp.log"
$Global:JsonLogFile = Join-Path $LogFolder "TeamsAttendanceRun_$TimeStamp.log.jsonl"
try { Start-Transcript -LiteralPath $TranscriptFile -Force -ErrorAction Stop | Out-Null; $Global:TranscriptOn = $true } catch { $Global:TranscriptOn = $false }

Write-Host ""
Write-Host "Configured:" -ForegroundColor Green
Write-Host "  Correlation ID : $Global:CorrelationId" -ForegroundColor Green
Write-Host ("  Window (PHT)   : {0} -> {1}" -f (Format-Pht $StartUtc), (Format-Pht $EndUtc)) -ForegroundColor Green
Write-Host ("  Window (UTC)   : {0} -> {1}" -f (Format-Utc $StartUtc), (Format-Utc $EndUtc)) -ForegroundColor DarkGray
Write-Host "  Attendance CSV : $AttCsv" -ForegroundColor Green
Write-Host "  Wide CSV       : $WideCsv" -ForegroundColor Green


# ============================================================
# PHASE 2 - Connect to Microsoft Graph (choose auth mode)
# ============================================================
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " PHASE 2: Connect to Microsoft Graph" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

foreach ($mod in @('Microsoft.Graph.Authentication','Microsoft.Graph.Calendar','Microsoft.Graph.CloudCommunications')) {
    if (-not (Ensure-Module -Name $mod)) { Write-Host "ERROR: $mod is required. Install-Module Microsoft.Graph -Scope CurrentUser -Force" -ForegroundColor Red; return }
}

Write-Host ""
$InputTenant = (Read-Host "Enter Tenant ID (GUID)").Trim().Trim('"').Trim("'")
$TenantId = (Sanitize-Path $InputTenant)
if ($TenantId -notmatch '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$') { Write-Host "ERROR: Tenant ID must be a GUID." -ForegroundColor Red; return }

Write-Host ""
Write-Host "Auth mode:" -ForegroundColor Cyan
Write-Host "  [1] Delegated interactive  - reads ONLY the signed-in user's own meetings" -ForegroundColor White
Write-Host "  [2] App-only (certificate) - all targeted organizers (needs app + Teams access policy)" -ForegroundColor White
Write-Host "  [3] App-only (client secret) - same as [2] with a secret" -ForegroundColor White
$AuthMode = (Read-Host "Enter 1, 2, or 3 (press Enter for 1)").Trim()
if ([string]::IsNullOrWhiteSpace($AuthMode)) { $AuthMode = '1' }

$Global:AppOnly = ($AuthMode -eq '2' -or $AuthMode -eq '3')
try {
    if ($AuthMode -eq '2') {
        $AppId = (Read-Host "  App (client) ID").Trim()
        $Thumb = (Read-Host "  Certificate thumbprint").Trim()
        Connect-MgGraph -TenantId $TenantId -ClientId $AppId -CertificateThumbprint $Thumb -NoWelcome -ErrorAction Stop
    }
    elseif ($AuthMode -eq '3') {
        $AppId  = (Read-Host "  App (client) ID").Trim()
        $Secret = Read-Host "  Client secret" -AsSecureString
        $cred = [System.Management.Automation.PSCredential]::new($AppId, $Secret)
        Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $cred -NoWelcome -ErrorAction Stop
    }
    else {
        $useDevice = (Read-Host "  Use device code? (Y/N, Enter=N)").Trim()
        $scopes = @('User.Read','Calendars.Read','OnlineMeetings.Read','OnlineMeetingArtifact.Read.All')
        if ($useDevice -match '^(y|yes)$') { Connect-MgGraph -TenantId $TenantId -Scopes $scopes -UseDeviceCode -NoWelcome -ErrorAction Stop }
        else { Connect-MgGraph -TenantId $TenantId -Scopes $scopes -NoWelcome -ErrorAction Stop }
    }
} catch { Write-Host "ERROR: Connect-MgGraph failed: $($_.Exception.Message)" -ForegroundColor Red; return }

$ctx = Get-MgContext
if (-not $ctx) { Write-Host "ERROR: No Graph context after connect." -ForegroundColor Red; return }
Write-Host ""
Write-Host "Connected." -ForegroundColor Green
Write-Host "  Account/App : $($ctx.Account)$($ctx.ClientId)" -ForegroundColor Green
Write-Host "  Tenant      : $($ctx.TenantId)" -ForegroundColor Green
Write-Host "  AuthType    : $($ctx.AuthType)" -ForegroundColor DarkGray


# ============================================================
# PHASE 3 - Verify Organizers CSV
# ============================================================
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " PHASE 3: Verify Organizers CSV" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

try { $Rows = Import-Csv -LiteralPath $CsvPath -ErrorAction Stop } catch { Write-Host "ERROR: Failed to read CSV: $($_.Exception.Message)" -ForegroundColor Red; return }
if ($null -eq $Rows -or @($Rows).Count -eq 0) { Write-Host "ERROR: CSV is empty." -ForegroundColor Red; return }
$cols = $Rows[0].PSObject.Properties.Name
$UpnCol = $null
foreach ($cand in @('upn','UPN','UserPrincipalName','userprincipalname','Upn')) { if ($cols -contains $cand) { $UpnCol = $cand; break } }
if (-not $UpnCol) { Write-Host "ERROR: CSV needs a 'upn' or 'UserPrincipalName' column. Found: $($cols -join ', ')" -ForegroundColor Red; return }

$Organizers = [System.Collections.Generic.List[string]]::new()
foreach ($r in $Rows) {
    $res = Sanitize-Upn -Value ([string]$r.$UpnCol)
    if ($res.Reason -ne 'Invalid UPN format' -and -not [string]::IsNullOrWhiteSpace($res.Clean) -and -not $Organizers.Contains($res.Clean)) { $Organizers.Add($res.Clean) }
}
if ($Organizers.Count -eq 0) { Write-Host "ERROR: No valid organizer UPNs." -ForegroundColor Red; return }

# Delegated mode can only read the signed-in user's own meetings.
if (-not $Global:AppOnly) {
    $me = [string]$ctx.Account
    Write-Host ""
    Write-Host "[NOTE] Delegated mode can read ONLY your own meetings ($me)." -ForegroundColor Yellow
    $selfOnly = $Organizers | Where-Object { $_ -eq $me.ToLower() }
    if (@($selfOnly).Count -eq 0) {
        Write-Host "       Your account isn't in the CSV; using your signed-in account as the only organizer." -ForegroundColor DarkYellow
        $Organizers = [System.Collections.Generic.List[string]]::new(); if ($me) { $Organizers.Add($me.ToLower()) }
    } else {
        $Organizers = [System.Collections.Generic.List[string]]::new(); $Organizers.Add($me.ToLower())
    }
}

Write-Host ""
Write-Host "Organizers to process: $($Organizers.Count)" -ForegroundColor Green
$Organizers | Select-Object -First 20 | ForEach-Object { Write-Host "  - $_" -ForegroundColor DarkGray }


# ============================================================
# PHASE 4 - Confirmation
# ============================================================
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " PHASE 4: Confirmation" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "About to READ Teams attendance (READ-ONLY):" -ForegroundColor Yellow
Write-Host "  Organizers   : $($Organizers.Count)" -ForegroundColor Yellow
Write-Host ("  Window (PHT) : {0} -> {1}" -f (Format-Pht $StartUtc), (Format-Pht $EndUtc)) -ForegroundColor Yellow
Write-Host "  Auth         : $(if($Global:AppOnly){'App-only'}else{'Delegated (self only)'})" -ForegroundColor Yellow
Write-Host ""
$Confirm = Read-Host "Proceed? (Y/N)"
if ($Confirm -notmatch '^(y|yes)$') { Write-Host "Cancelled." -ForegroundColor Red; try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}; if ($Global:TranscriptOn) { try { Stop-Transcript | Out-Null } catch {} }; return }


# ============================================================
# PHASE 5 - Extract Attendance
# ============================================================
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " PHASE 5: Extract Attendance" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

$RunStart = Get-Date
$AttRows  = [System.Collections.Generic.List[object]]::new()
$WideTmp  = [System.Collections.Generic.List[object]]::new()
$Counters = @{ Meetings=0; Reports=0; Records=0; Organizers=0; NoMeeting=0; OutOfWindow=0 }

# ISO (UTC) strings for CalendarView.
$sIso = ([datetime]$StartUtc).ToString('yyyy-MM-ddTHH:mm:ssZ')
$eIso = ([datetime]$EndUtc).ToString('yyyy-MM-ddTHH:mm:ssZ')

# The Teams onlineMeetings API needs the user's OBJECT ID (GUID), not the UPN.
# Resolve UPN -> id (or accept an id already in the CSV). Cached.
$UserIdCache = @{}
function Resolve-UserId {
    param([string]$UpnOrId)
    if ([string]::IsNullOrWhiteSpace($UpnOrId)) { return $null }
    if ($UpnOrId -match '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$') { return $UpnOrId }  # already an object id
    $k = $UpnOrId.ToLower()
    if ($UserIdCache.ContainsKey($k)) { return $UserIdCache[$k] }
    $id = $null
    try {
        $u   = [uri]::EscapeDataString($UpnOrId)
        $uri = "/v1.0/users/$($u)?`$select=id"
        $resp = Invoke-WithRetry { Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop }
        if ($resp) {
            if ($resp -is [System.Collections.IDictionary]) { if ($resp.Contains('id')) { $id = [string]$resp['id'] } }
            else { $id = [string](Get-Prop $resp @('id','Id')) }
        }
    } catch { $id = $null }
    $UserIdCache[$k] = $id
    return $id
}

# Delegated: resolve the signed-in user's own id via /me (needs only User.Read).
$SelfId = $null
if (-not $Global:AppOnly) {
    try { $meR = Invoke-MgGraphRequest -Method GET -Uri "/v1.0/me?`$select=id" -ErrorAction Stop; if ($meR -is [System.Collections.IDictionary] -and $meR.Contains('id')) { $SelfId = [string]$meR['id'] } } catch { }
}

foreach ($org in $Organizers) {
    $Counters.Organizers++
    Write-Log "Organizer $($Counters.Organizers)/$($Organizers.Count): $org" 'INFO'

    # Resolve to an object id (GUID) - required by the Teams onlineMeetings API.
    $orgId = if (-not $Global:AppOnly -and $SelfId) { $SelfId } else { Resolve-UserId $org }
    if ([string]::IsNullOrWhiteSpace($orgId)) {
        Write-Log "  Could not resolve '$org' to an object id. Put the user's OBJECT ID in the CSV, or (app-only) grant User.ReadBasic.All so UPNs can be translated. Skipping." 'WARN'
        continue
    }

    # 1) Calendar view -> Teams online meetings in range
    $events = @()
    try {
        $events = Invoke-WithRetry {
            Get-MgUserCalendarView -UserId $orgId -StartDateTime $sIso -EndDateTime $eIso -All `
                -Property 'id,subject,start,end,isOnlineMeeting,onlineMeeting,organizer' -ErrorAction Stop
        }
    } catch { Write-Log "  Calendar read failed for ${org}: $($_.Exception.Message)" 'WARN'; continue }

    $onlineEvents = @($events | Where-Object { $_.IsOnlineMeeting -and $_.OnlineMeeting -and $_.OnlineMeeting.JoinUrl })
    Write-Log "  $($onlineEvents.Count) Teams meeting(s) on calendar." 'INFO'

    $seenJoinUrls = [System.Collections.Generic.HashSet[string]]::new()
    $evIdx = 0
    foreach ($ev in $onlineEvents) {
        $joinUrl = [string]$ev.OnlineMeeting.JoinUrl
        if (-not $seenJoinUrls.Add($joinUrl)) { continue }   # avoid duplicate recurring instances of same series
        $subject = Get-Str $ev @('Subject')
        $evIdx++
        Write-Log "  Meeting $evIdx/$($onlineEvents.Count): resolving '$subject'..." 'INFO'
        # Host display name from the calendar event organizer (no User.Read.All needed).
        $orgName = ''
        try { if ($ev.Organizer -and $ev.Organizer.EmailAddress) { $orgName = Get-Str $ev.Organizer.EmailAddress @('Name') } } catch { }

        # 2) Resolve JoinWebUrl -> onlineMeeting
        $om = $null
        try {
            $esc = $joinUrl.Replace("'","''")
            $om = Invoke-WithRetry { Get-MgUserOnlineMeeting -UserId $orgId -Filter "JoinWebUrl eq '$esc'" -ErrorAction Stop } | Select-Object -First 1
        } catch { Write-Log "  onlineMeeting lookup failed: $($_.Exception.Message)" 'WARN'; continue }
        if (-not $om) { $Counters.NoMeeting++; continue }
        $omId = [string]$om.Id
        $Counters.Meetings++

        # 3) Attendance reports for this meeting
        $reports = @()
        try { $reports = Invoke-WithRetry { Get-MgUserOnlineMeetingAttendanceReport -UserId $orgId -OnlineMeetingId $omId -All -ErrorAction Stop } }
        catch { Write-Log "  attendanceReports failed: $($_.Exception.Message)" 'WARN'; continue }

        foreach ($rep in @($reports)) {
            $repId = [string]$rep.Id
            $mStart = ConvertTo-Utc (Get-Prop $rep @('MeetingStartDateTime'))
            $mEnd   = ConvertTo-Utc (Get-Prop $rep @('MeetingEndDateTime'))

            # A recurring meeting is ONE onlineMeeting with MANY reports (its whole history).
            # Keep only reports whose session falls inside the requested window.
            $repWhen = if ($null -ne $mStart) { $mStart } elseif ($null -ne $mEnd) { $mEnd } else { $null }
            if ($null -ne $repWhen -and ($repWhen -lt $StartUtc -or $repWhen -ge $EndUtc)) { $Counters.OutOfWindow++; continue }
            $Counters.Reports++

            $records = @()
            try { $records = Invoke-WithRetry { Get-MgUserOnlineMeetingAttendanceReportAttendanceRecord -UserId $orgId -OnlineMeetingId $omId -MeetingAttendanceReportId $repId -All -ErrorAction Stop } }
            catch { Write-Log "  attendanceRecords failed: $($_.Exception.Message)" 'WARN'; continue }
            Write-Log "    report: $(@($records).Count) attendance record(s)." 'INFO'

            $partList = [System.Collections.Generic.List[string]]::new()
            foreach ($rec in @($records)) {
                $Counters.Records++
                try { ($rec | ConvertTo-Json -Depth 6 -Compress) | Add-Content -LiteralPath $RawJsonl -Encoding UTF8 } catch { }

                $email = Get-Str $rec @('EmailAddress')
                $name  = ''
                try { if ($rec.Identity) { $name = Get-Str $rec.Identity @('DisplayName') } } catch { }
                $role  = Get-Str $rec @('Role')
                $totalSec = 0; try { $totalSec = [int](Get-Prop $rec @('TotalAttendanceInSeconds') 0) } catch { }

                # First join / last leave from intervals
                $firstJoin = $null; $lastLeave = $null
                try {
                    foreach ($iv in @($rec.AttendanceIntervals)) {
                        $j = ConvertTo-Utc (Get-Prop $iv @('JoinDateTime'))
                        $l = ConvertTo-Utc (Get-Prop $iv @('LeaveDateTime'))
                        if ($j -and (-not $firstJoin -or $j -lt $firstJoin)) { $firstJoin = $j }
                        if ($l -and (-not $lastLeave -or $l -gt $lastLeave)) { $lastLeave = $l }
                    }
                } catch { }

                $idKey = if ($email) { $email } else { $name }
                if ($idKey -and -not ($partList -contains $idKey)) { $partList.Add($idKey) }

                $AttRows.Add([PSCustomObject]@{
                    CorrelationId   = $Global:CorrelationId
                    Organizer       = $org
                    Subject         = $subject
                    MeetingStartPHT = (Format-Pht $mStart)
                    MeetingEndPHT   = (Format-Pht $mEnd)
                    ParticipantName = $name
                    ParticipantEmail= $email
                    Role            = $role
                    FirstJoinPHT    = (Format-Pht $firstJoin)
                    LastLeavePHT    = (Format-Pht $lastLeave)
                    AttendanceMin   = [math]::Round($totalSec/60,1)
                    OnlineMeetingId = $omId
                })
            }

            $WideTmp.Add([PSCustomObject]@{
                Organizer=$org; HostDisp=$orgName; Subject=$subject; Start=(Format-Pht $mStart); End=(Format-Pht $mEnd); MeetingId=$omId; Parts=$partList
            })
        }
    }
}

# --- Build wide rows ---
$maxParts = 0
foreach ($w in $WideTmp) { if (@($w.Parts).Count -gt $maxParts) { $maxParts = @($w.Parts).Count } }
$WideRows = [System.Collections.Generic.List[object]]::new()
foreach ($w in $WideTmp) {
    $row = [ordered]@{
        'Host UPN'          = $w.Organizer
        'Display Name'      = $w.HostDisp
        'Subject'           = $w.Subject
        'Start (PHT)'       = $w.Start
        'End (PHT)'         = $w.End
        'MeetingId'         = $w.MeetingId
        'Participant Count' = @($w.Parts).Count
    }
    for ($i=0; $i -lt $maxParts; $i++) { $row["Participant $($i+1)"] = if ($i -lt @($w.Parts).Count) { $w.Parts[$i] } else { '' } }
    $WideRows.Add([PSCustomObject]$row)
}

# --- Export ---
try {
    if ($AttRows.Count -gt 0) { $AttRows | Sort-Object Organizer, MeetingStartPHT, ParticipantName | Export-Csv -LiteralPath $AttCsv -NoTypeInformation -Encoding UTF8 }
    else { 'CorrelationId,Organizer,Subject,MeetingStartPHT,MeetingEndPHT,ParticipantName,ParticipantEmail,Role,FirstJoinPHT,LastLeavePHT,AttendanceMin,OnlineMeetingId' | Set-Content -LiteralPath $AttCsv -Encoding UTF8 }
    if ($WideRows.Count -gt 0) { $WideRows | Sort-Object 'Host UPN','Start (PHT)' | Export-Csv -LiteralPath $WideCsv -NoTypeInformation -Encoding UTF8 }
    else { 'Host UPN,Display Name,Subject,Start (PHT),End (PHT),MeetingId,Participant Count' | Set-Content -LiteralPath $WideCsv -Encoding UTF8 }
    Write-Log "Export complete." 'SUCCESS'
} catch { Write-Host "ERROR: Export failed: $($_.Exception.Message)" -ForegroundColor Red }

$dur = (Get-Date) - $RunStart
Write-Host ""
Write-Host "=== EXECUTION SUMMARY ===" -ForegroundColor Cyan
Write-Host (" Correlation ID   : {0}" -f $Global:CorrelationId)
Write-Host (" Window (PHT)     : {0} -> {1}" -f (Format-Pht $StartUtc), (Format-Pht $EndUtc))
Write-Host (" Organizers       : {0}" -f $Counters.Organizers)
Write-Host (" Meetings matched : {0}" -f $Counters.Meetings)
Write-Host (" No onlineMeeting : {0}" -f $Counters.NoMeeting)
Write-Host (" Attendance reports:{0}  (in-window)" -f $Counters.Reports)
Write-Host (" Reports skipped   : {0}  (recurring-series sessions outside your date window)" -f $Counters.OutOfWindow) -ForegroundColor DarkGray
Write-Host (" Attendance rows  : {0}" -f $AttRows.Count) -ForegroundColor Green
Write-Host (" Duration         : {0:hh\:mm\:ss}" -f $dur)
Write-Host (" Attendance CSV   : {0}" -f $AttCsv)
Write-Host (" Wide CSV         : {0}" -f $WideCsv)
Write-Host " CSV times are Manila (PHT, UTC+8). Raw JSONL keeps UTC." -ForegroundColor DarkGray
Write-Host "=========================" -ForegroundColor Cyan
if ($AttRows.Count -eq 0) {
    Write-Host "[NOTE] 0 attendance rows. Common causes: meetings had no attendance report" -ForegroundColor Yellow
    Write-Host "       (didn't occur, or reports disabled by Teams policy); app access policy not" -ForegroundColor Yellow
    Write-Host "       granted (app-only); reports aged out; or 'Meet now' calls with no calendar event." -ForegroundColor Yellow
}

try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
if ((Get-Variable -Name TranscriptOn -Scope Global -ErrorAction SilentlyContinue) -and $Global:TranscriptOn) { try { Stop-Transcript | Out-Null } catch { } }


<#
============================================================
 SETUP NOTES
============================================================
DELEGATED (mode 1): no setup. You can only read YOUR OWN organized meetings.
  Scopes consented at sign-in: User.Read, Calendars.Read, OnlineMeetings.Read,
  OnlineMeetingArtifact.Read.All. (Self id resolved via /me.)

APP-ONLY (mode 2/3) - LEAST-PRIVILEGE recipe for "these organizers only":
  0. Create ONE mail-enabled security group, e.g. "Teams-Attendance-Organizers", and add the
     teacher accounts. This group is your control surface (add/remove a teacher here later).
  1. Entra admin center > App registrations > New registration. Note App (client) ID + Tenant ID.
  2. API permissions > Microsoft Graph > Application permissions, add (all read-only):
        Calendars.Read, OnlineMeetings.Read.All, OnlineMeetingArtifact.Read.All
     Optional, only if the CSV holds UPNs (to translate UPN -> object id):  User.ReadBasic.All
     -> to avoid even that, put each organizer's OBJECT ID in the CSV instead of the UPN.
     Then "Grant admin consent".
  3. Credentials: upload a certificate (mode 2) or create a client secret (mode 3).
  4. SCOPE MEETINGS to the group (Teams PowerShell): without this, app-only meeting calls 403.
        New-CsApplicationAccessPolicy -Identity "TeamsAttendance" -AppIds "<APP_ID>" -Description "Attendance export"
        Grant-CsApplicationAccessPolicy -Group "Teams-Attendance-Organizers@<domain>" -PolicyName "TeamsAttendance" -Rank 1
        # (or -Identity <organizerUPN> per teacher; or -Global for everyone - broadest)
  5. SCOPE CALENDARS to the same group (Exchange Online PowerShell) so Calendars.Read is not tenant-wide:
        New-ApplicationAccessPolicy -AppId "<APP_ID>" -PolicyScopeGroupId "Teams-Attendance-Organizers@<domain>" `
            -AccessRight RestrictAccess -Description "Attendance export - calendars"
  Notes: policy changes take up to 30 min to propagate. Cert auth (mode 2) needs the cert in the
  runner's CurrentUser\My store (enter its thumbprint at the prompt).
============================================================
#>
