#Requires -Version 7.0
#Requires -Modules Microsoft.Graph.Users, Microsoft.Graph.Users.Actions, Microsoft.Graph.Identity.DirectoryManagement

<#
.SYNOPSIS
    Bulk reset Entra ID user passwords with optional cryptographic random password generation.

.DESCRIPTION
    Phased (5-block) copy-paste script for resetting cloud-only Entra ID user passwords.
    Reads a CSV with header 'upn' and an OPTIONAL 'password' column.

    Password sources (chosen at runtime in Phase 4):
      [1] CSV only  - use the password column for every row (v2.x behaviour)
      [2] Generate  - ignore the CSV password column, generate for every row
      [3] Hybrid    - generate ONLY where the password cell is blank/missing  <-- default

    Password styles:
      [1] Readable  - Copper-Bright-Falcon-4827!   (easy to dictate / print / type on a tablet)
      [2] Random    - Kx7#mQp4$vRt9w               (maximum entropy)

    Output separation (important):
      *_CREDENTIALS_*.csv  -> cleartext passwords for distribution. WRITTEN BEFORE ANY RESET RUNS.
      *_Log_*.csv          -> audit trail. NO cleartext. Carries source, length and a run-salted
                              fingerprint only, so it is safe to attach to a ticket or SharePoint.

.AUTHOR         Generated with Claude for Rakso CT Education IT.
.VERSION        3.7.0
.DATE           2026-08-19
.REQUIREMENTS   PowerShell 7+; Microsoft.Graph SDK
                (Users, Users.Actions, Identity.DirectoryManagement).
.PERMISSIONS    User-PasswordProfile.ReadWrite.All   <-- MANDATORY for password reset
                User.RevokeSessions.All              <-- only for optional force sign-out
                                                         (User.ReadWrite.All also satisfies it)
                User.ReadWrite.All
                Directory.Read.All
                RoleManagement.Read.Directory
                Directory role: User Administrator (members) or
                                Privileged Authentication Administrator (to touch admins).
                A PIM-eligible role must be ACTIVATED before running.
.SAFETY         Dry-run mode, scope pre-verification, operator role check, cloud-only guard,
                privileged-account gate, large-batch gate, 403 circuit breaker,
                credentials pre-write (crash cannot orphan a password), masked console output.
.CHANGELOG      v2.1.0 - Sanitization, preflight, retry, manifest.
                v3.0.0 - Cryptographic password generation (readable + random styles),
                         hybrid per-row password source, split credentials/audit outputs with
                         pre-write, batch uniqueness, masked console, retry-safe re-run CSV,
                         UserObjectId populated, correlation ID, Set-StrictMode.
                v3.1.0 - Added User-PasswordProfile.ReadWrite.All (Graph split password reset
                         out of User.ReadWrite.All on 2024-12-23; without it EVERY row 403s).
                         Phase 2 now verifies granted scopes and the operator's active roles
                         and aborts before writing anything. Phase 5 compresses the Graph SDK
                         error dump to one line + request-id, and a circuit breaker halts the
                         run after 5 consecutive 403s instead of burning every row. Summary
                         refuses to green-light distributing an unapplied credentials file.
                v3.2.0 - Configurable progress heartbeat (0 = off) applied to both pre-flight
                         and execution. Optional force sign-out per successful reset via
                         Revoke-MgUserSignInSession (Microsoft.Graph.Users.Actions), gated on
                         scope availability, skipped for the operator's own account, and
                         auto-disabled after 3 consecutive 403s so resets keep running.
                         New audit columns: SessionRevokeStatus, SessionRevokeMessage.
                v3.3.0 - Sign-in default changed from device code to interactive popup: Graph
                         SDK 2.34+ moved Azure.Identity to the WAM broker and broke device-code
                         token caching, so -UseDeviceCode connects then fails every cmdlet with
                         'DeviceCodeCredential authentication failed: Object reference not set
                         to an instance of an object' (SDK issue #3495). Phase 2 now smoke-tests
                         the token with a live read, diagnoses that exact fault, and offers an
                         in-place interactive reconnect. Pre-flight distinguishes auth failures
                         from 'user not found' and aborts instead of mislabelling every row.
                v3.4.0 - Performance and visibility. The v3.3.0 operator role check enumerated
                         every activated directory role and then paged every member of each
                         (O(roles) requests, no output) - replaced with a single
                         Get-MgUserMemberOf call. The privileged-account scan now prints
                         per-role progress, member counts and elapsed time, retries through
                         throttling, and can be switched OFF in Run options. User resolution
                         reports total and per-user timing. No silent multi-second blocks left.
                v3.5.0 - Scale. Pre-flight used one Graph request per row, so a 23,145-row CSV
                         meant 23,145 sequential lookups (1-2 hours before any write). Batches
                         of 200+ now page the directory once at 999/page (~24 requests) into a
                         UPN-keyed index and match in memory, with per-page progress; falls
                         back to per-user lookups if the bulk read fails. Audit-log flushing
                         changed from full rewrite to append-only - the old path was O(n^2) and
                         at 23k rows would have rewritten ~26 million rows over a run.
                         Added a pre-commit duration estimate, and a live rows/s + ETA in the
                         heartbeat with a warning if heartbeat is off on a long run.
                v3.6.0 - Throughput and restartability.
                         * Batched write mode: 20 PATCHes per Graph $batch call, executed in
                           parallel server-side. Entra allows ~3,000 writes / 2.5 min per app
                           per tenant (~20/sec); sequential writes only reached 2-3/sec, so
                           this is a ~6-8x cut in wall clock. Per-sub-request status parsing;
                           only throttled sub-requests are retried.
                         * Throttle governor: paces proactively to ~18 writes/sec, backs off
                           on 429 (honouring Retry-After) and recovers after clean batches.
                         * CSV row range: process rows N..M of the SAME file - no manual
                           splitting, original row numbers preserved across slices.
                         * Resume: point at a previous audit log and every UPN already marked
                           Success is skipped before any lookup or password generation.
                         * Audit schema unified via New-AuditRow; new WriteMode column. In
                           batched mode the log is in EXECUTION order - sort by Row for CSV
                           order (recorded in the manifest).
                v3.7.0 - Portability. Removed the hardcoded operator-specific default
                         input CSV and log folder (they pointed at one admin's OneDrive
                         and silently broke for everyone else). Defaults are now derived
                         at runtime from the script's own folder and remain overridable
                         at the prompt, so the file contains nothing machine-specific
                         and is safe to commit to source control. Tenant ID was already
                         prompt-only and is unchanged.

.NOTES
    Fail-safes:
      - Invisible-char stripping (U+202A, NBSP, ZWSP, BOM, etc.)
      - UPN sanitization + NFD Unicode normalization + 113-char Entra limit
      - Duplicate UPN detection (first row wins)
      - Password policy pre-check (length / all-numeric / matches UPN local part)
      - Dry-run mode (validate + generate without writing to the tenant)
      - Cloud-only guard (blocks hybrid / OnPremisesSyncEnabled users)
      - Privileged account detection (GA / PA / UA / PRA gate)
      - Retry with exponential backoff + Retry-After header support
      - Credentials file written and verified BEFORE Phase 5 executes
      - Log flush every 10 rows (crash-safe), heartbeat every 25 rows

    How to use:
      Copy and paste each phase one at a time, or paste the whole file at once.
#>

Set-StrictMode -Version Latest

$ScriptVersion   = "3.7.0"
$CorrelationId   = [guid]::NewGuid().ToString('N').Substring(0,12)


# ============================================================
# PHASE 0 - HELPER FUNCTIONS (loaded once, used across all phases)
# ============================================================

# ---------- String / path hygiene ----------

# Strip invisible Unicode formatting chars from pasted strings
function Remove-InvisibleChars {
    param([string]$s)
    if ([string]::IsNullOrEmpty($s)) { return $s }
    return ($s -replace '[\u200B-\u200F\u202A-\u202E\u2066-\u2069\uFEFF]', '')
}

# Sanitize a UPN: strip invisible/whitespace chars, NFD-normalize accents, lowercase, validate.
# (Renamed from Sanitize-Upn in v2.1 - 'Sanitize' is not an approved PowerShell verb.)
function Convert-UpnToCanonical {
    param([AllowEmptyString()][string]$Raw)

    if ([string]::IsNullOrEmpty($Raw)) {
        return [PSCustomObject]@{ Clean = ''; Changed = $false; Valid = $false; Reason = 'Empty' }
    }

    # 1. Strip invisible formatting + control chars
    $s = $Raw -replace '[\u200B-\u200F\u202A-\u202E\u2066-\u2069\uFEFF\u0000-\u001F\u007F]', ''

    # 2. Strip Unicode space variants (NBSP, narrow NBSP, hair spaces, etc.)
    $s = $s -replace '[\u00A0\u1680\u2000-\u200A\u202F\u205F\u3000]', ''

    # 3. Trim regular whitespace + quotes
    $s = $s.Trim().Trim('"').Trim("'").Trim()

    # 4. NFD normalize (e-acute -> e + combining acute) then drop combining marks -> ASCII fold
    $normalized = $s.Normalize([Text.NormalizationForm]::FormD)
    $sb = [Text.StringBuilder]::new()
    foreach ($ch in $normalized.ToCharArray()) {
        if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$sb.Append($ch)
        }
    }
    $s = $sb.ToString().ToLowerInvariant()

    # 5. Validate
    $changed = ($s -ne $Raw)
    $valid   = $true
    $reason  = 'OK'
    if ($s -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') { $valid = $false; $reason = 'Invalid UPN format' }
    elseif ($s.Length -gt 113)                     { $valid = $false; $reason = 'Exceeds 113-char Entra limit' }

    return [PSCustomObject]@{
        Clean   = $s
        Changed = $changed
        Valid   = $valid
        Reason  = $reason
    }
}

function Convert-PathSafe {
    param([AllowEmptyString()][string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
    return (Remove-InvisibleChars $Path).Trim().Trim('"').Trim("'").Trim()
}

# ---------- Cryptographic randomness ----------
# NOTE: Get-Random is a SEEDABLE, NON-cryptographic PRNG and must never be used for
# credentials. Everything below draws from RandomNumberGenerator (CSPRNG, unbiased).

function Get-CryptoInt {
    param([Parameter(Mandatory)][int]$MaxExclusive)
    if ($MaxExclusive -le 0) { throw "Get-CryptoInt: MaxExclusive must be greater than 0." }
    return [System.Security.Cryptography.RandomNumberGenerator]::GetInt32($MaxExclusive)
}

# Pick one element from a string (treated as a char pool) or from an array
function Get-CryptoPick {
    param([Parameter(Mandatory)]$Pool)
    if ($Pool -is [string]) {
        $arr = $Pool.ToCharArray()
    } else {
        $arr = @($Pool)
    }
    if ($arr.Length -eq 0) { throw "Get-CryptoPick: pool is empty." }
    return $arr[(Get-CryptoInt -MaxExclusive $arr.Length)]
}

# Fisher-Yates shuffle using the CSPRNG.
# Without this, a password built as Upper+lower+digit+symbol+filler leaks its own structure.
function Get-CryptoShuffledString {
    param([Parameter(Mandatory)][char[]]$Chars)
    for ($i = $Chars.Length - 1; $i -gt 0; $i--) {
        $j = Get-CryptoInt -MaxExclusive ($i + 1)
        $t = $Chars[$i]; $Chars[$i] = $Chars[$j]; $Chars[$j] = $t
    }
    return (-join $Chars)
}

# ---------- Password character pools ----------
# Visually ambiguous characters are excluded on purpose: these passwords get dictated
# over the phone, printed on handouts, and typed by students on tablets.
#   Removed uppercase : I O S Z
#   Removed lowercase : l o
#   Removed digits    : 0 1
#   Symbols limited to a CSV-safe, shell-safe, regional-keyboard-safe set.
$PwdPoolUpper  = 'ABCDEFGHJKLMNPQRTUVWXY'      # 22
$PwdPoolLower  = 'abcdefghijkmnpqrstuvwxyz'    # 24
$PwdPoolDigit  = '23456789'                    # 8
$PwdPoolSymbol = '!#$%*+-=?@'                  # 10   -> combined pool = 64 chars = 6.00 bits/char
$PwdPoolTailSymbol = '!#$%*+=?@'               # 9    (no '-' : that is the readable separator)

# ---------- Embedded word lists (readable style) ----------
# Curated: 4-10 letters, unambiguous spelling, no profanity, no words on Microsoft's
# global banned-password list, no overlap between the two lists. Offline by design.
$PwdWordsA = @(
    'Amber','Arctic','Autumn','Azure','Bold','Brave','Bright','Brisk','Bronze','Calm',
    'Candid','Chief','Civic','Clear','Clever','Coastal','Cobalt','Cool','Copper','Coral',
    'Cosmic','Crimson','Crisp','Curious','Daring','Deep','Dense','Diamond','Eager','Early',
    'Earthy','Elder','Electric','Elegant','Emerald','Epic','Even','Fair','Famous','Fancy',
    'Fast','Fearless','Fine','Firm','First','Fluent','Fond','Formal','Forward','Fresh',
    'Frosty','Gallant','Gentle','Giant','Gifted','Glad','Global','Golden','Grand','Granite',
    'Green','Happy','Hardy','Hearty','Helpful','Hidden','Honest','Humble','Indigo','Inland',
    'Ivory','Jade','Jolly','Joyful','Keen','Kind','Lively','Loyal','Lucid','Lucky',
    'Magenta','Marble','Mellow','Merry','Mighty','Mint','Modern','Native','Neat','Noble',
    'Northern','Olive','Onyx','Opal','Orange','Pacific','Patient','Peach','Pearl','Plain',
    'Polar','Polite','Prime','Proud','Purple','Quartz','Quick','Quiet','Rapid','Ready',
    'Regal','Robust','Royal','Ruby','Rugged','Rustic','Sandy','Sapphire','Scarlet','Serene',
    'Sharp','Silver','Simple','Sincere','Slate','Sleek','Smart','Smooth','Snowy','Solar',
    'Solid','Sonic','Spring','Stable','Steady','Sterling','Stormy','Strong','Sturdy','Summer',
    'Sunny','Swift','Teal','Tender','Thrifty','Tidal','Timely','Topaz','Tranquil','True',
    'Trusty','Turquoise','Upbeat','Urban','Valiant','Velvet','Vibrant','Violet','Vivid','Warm',
    'Watchful','Western','Whole','Wild','Winter','Wise','Witty','Woven','Yellow','Zesty'
)
$PwdWordsB = @(
    'Anchor','Antler','Arrow','Aspen','Badger','Bamboo','Basin','Beacon','Beaver','Birch',
    'Bison','Blossom','Boulder','Branch','Breeze','Bridge','Brook','Buffalo','Cactus','Camel',
    'Canyon','Cardinal','Cavern','Cedar','Cliff','Comet','Compass','Condor','Cottage','Cougar',
    'Crane','Crater','Creek','Crest','Cypress','Dolphin','Eagle','Ember','Falcon','Fern',
    'Fjord','Forest','Fountain','Galaxy','Garden','Gazelle','Geyser','Glacier','Glade','Grotto',
    'Grove','Harbor','Harvest','Hawk','Heron','Hickory','Hollow','Horizon','Iceberg','Island',
    'Jaguar','Jasmine','Juniper','Kestrel','Lagoon','Lantern','Laurel','Ledge','Leopard','Lily',
    'Lotus','Magnolia','Mammoth','Maple','Marlin','Meadow','Mesa','Meteor','Mongoose','Moose',
    'Mountain','Nebula','Nectar','Oasis','Ocean','Ocelot','Orbit','Orchid','Osprey','Otter',
    'Palm','Panda','Panther','Parrot','Pebble','Pelican','Penguin','Pillar','Pine','Planet',
    'Plateau','Poppy','Prairie','Puffin','Quarry','Quill','Rabbit','Raccoon','Rainbow','Rapids',
    'Raven','Reef','Ridge','River','Robin','Rocket','Sable','Salmon','Sequoia','Shelter',
    'Sparrow','Spruce','Stallion','Starling','Stream','Summit','Sunrise','Swallow','Sycamore','Thistle',
    'Thunder','Tiger','Timber','Toucan','Trail','Tulip','Tundra','Turtle','Valley','Vessel',
    'Village','Vine','Walnut','Walrus','Warbler','Waterfall','Willow','Wombat','Woodland','Zebra'
)

# ---------- Password generators ----------

# Style 1: random string. Guarantees one char from each of the 4 classes, then shuffles.
function New-RandomStringPassword {
    param([int]$Length = 14)
    if ($Length -lt 12) { $Length = 12 }
    if ($Length -gt 64) { $Length = 64 }
    $all   = $PwdPoolUpper + $PwdPoolLower + $PwdPoolDigit + $PwdPoolSymbol
    $chars = [System.Collections.Generic.List[char]]::new()
    [void]$chars.Add((Get-CryptoPick -Pool $PwdPoolUpper))
    [void]$chars.Add((Get-CryptoPick -Pool $PwdPoolLower))
    [void]$chars.Add((Get-CryptoPick -Pool $PwdPoolDigit))
    [void]$chars.Add((Get-CryptoPick -Pool $PwdPoolSymbol))
    while ($chars.Count -lt $Length) { [void]$chars.Add((Get-CryptoPick -Pool $all)) }
    return (Get-CryptoShuffledString -Chars $chars.ToArray())
}

# Style 2: readable pattern, e.g. Copper-Bright-Falcon-4827!
# Satisfies all 4 Entra character classes (upper, lower, digit, symbol).
function New-ReadablePassword {
    param(
        [int]$WordCount  = 3,
        [int]$DigitCount = 4,
        [string]$Separator = '-'
    )
    if ($WordCount  -lt 2) { $WordCount  = 2 }
    if ($WordCount  -gt 4) { $WordCount  = 4 }
    if ($DigitCount -lt 2) { $DigitCount = 2 }
    if ($DigitCount -gt 6) { $DigitCount = 6 }

    $parts = [System.Collections.Generic.List[string]]::new()

    # (WordCount - 1) modifier words, kept distinct so we never emit 'Copper-Copper-Falcon'
    $usedA = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    for ($i = 1; $i -lt $WordCount; $i++) {
        for ($try = 0; $try -lt 50; $try++) {
            $w = [string](Get-CryptoPick -Pool $PwdWordsA)
            if ($usedA.Add($w)) { [void]$parts.Add($w); break }
        }
    }
    [void]$parts.Add([string](Get-CryptoPick -Pool $PwdWordsB))

    $digits = -join (1..$DigitCount | ForEach-Object { [string](Get-CryptoInt -MaxExclusive 10) })
    [void]$parts.Add($digits)

    return (($parts -join $Separator) + [string](Get-CryptoPick -Pool $PwdPoolTailSymbol))
}

# Honest entropy accounting, printed to the operator before they commit.
function Get-PasswordEntropyBits {
    param(
        [Parameter(Mandatory)][ValidateSet('Readable','Random')][string]$Style,
        [int]$Length = 14,
        [int]$WordCount = 3,
        [int]$DigitCount = 4
    )
    if ($Style -eq 'Random') {
        $poolSize = ($PwdPoolUpper + $PwdPoolLower + $PwdPoolDigit + $PwdPoolSymbol).Length
        return [Math]::Round($Length * [Math]::Log($poolSize, 2), 1)
    }
    $bits  = [Math]::Log($PwdWordsA.Count, 2) * ($WordCount - 1)
    $bits += [Math]::Log($PwdWordsB.Count, 2)
    $bits += [Math]::Log(10, 2) * $DigitCount
    $bits += [Math]::Log($PwdPoolTailSymbol.Length, 2)
    return [Math]::Round($bits, 1)
}

# Generate a password that has not already been issued in this batch.
function New-UniquePassword {
    param(
        # AllowEmptyCollection is REQUIRED: on the first row the set is empty, and PowerShell's
        # Mandatory validation rejects an empty collection outright.
        [Parameter(Mandatory)][AllowEmptyCollection()][System.Collections.Generic.HashSet[string]]$UsedSet,
        [Parameter(Mandatory)][ValidateSet('Readable','Random')][string]$Style,
        [int]$Length = 14,
        [int]$WordCount = 3,
        [int]$DigitCount = 4
    )
    for ($attempt = 1; $attempt -le 100; $attempt++) {
        $p = if ($Style -eq 'Random') {
                New-RandomStringPassword -Length $Length
             } else {
                New-ReadablePassword -WordCount $WordCount -DigitCount $DigitCount
             }
        if ($UsedSet.Add($p)) { return $p }
    }
    throw "New-UniquePassword: could not produce a unique password after 100 attempts."
}

# ---------- Secret handling ----------

# Never print a full generated password to the console (screenshare / scrollback risk).
function Format-MaskedSecret {
    param([AllowEmptyString()][string]$Secret)
    if ([string]::IsNullOrEmpty($Secret)) { return '<none>' }
    $len = $Secret.Length
    if ($len -le 4) { return ('*' * $len) }
    return ($Secret.Substring(0,2) + ('*' * ($len - 2)) + " ($len ch)")
}

# Run-salted, truncated fingerprint. Salting with the correlation ID means the audit log
# cannot be dictionary-attacked back to the plaintext, but rows within one run stay comparable.
function Get-SecretFingerprint {
    param([AllowEmptyString()][string]$Secret, [Parameter(Mandatory)][string]$Salt)
    if ([string]::IsNullOrEmpty($Secret)) { return '' }
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [Text.Encoding]::UTF8.GetBytes("$Salt|$Secret")
        $hex   = ([BitConverter]::ToString($sha.ComputeHash($bytes)) -replace '-', '').ToLowerInvariant()
        return $hex.Substring(0,16)
    }
    finally { $sha.Dispose() }
}

# Shared policy pre-check. Applied uniformly to CSV-supplied AND generated passwords
# so there is exactly one definition of "acceptable" in the script.
function Test-PasswordPolicy {
    param([AllowEmptyString()][string]$Password, [AllowEmptyString()][string]$Upn)
    $issues = @()
    if ([string]::IsNullOrWhiteSpace($Password)) { return 'Empty' }
    if ($Password.Length -lt 8)   { $issues += 'Too short (<8 chars)' }
    if ($Password.Length -gt 256) { $issues += 'Exceeds 256-char Entra limit' }
    if ($Password -match '^\d+$') { $issues += 'All numeric (weak)' }
    if ($Upn) {
        $local = ($Upn -split '@')[0]
        if ($local -and $Password.ToLowerInvariant() -eq $local.ToLowerInvariant()) {
            $issues += 'Matches UPN local part'
        }
    }
    # Entra requires 3 of 4 character classes for passwords under 256 chars
    $classes = 0
    if ($Password -cmatch '[a-z]')        { $classes++ }
    if ($Password -cmatch '[A-Z]')        { $classes++ }
    if ($Password -match  '\d')           { $classes++ }
    if ($Password -match  '[^a-zA-Z0-9]') { $classes++ }
    if ($classes -lt 3) { $issues += "Only $classes/4 character classes (Entra needs 3)" }
    return ($issues -join '; ')
}

# ---------- Graph plumbing ----------

# Graph call with exponential backoff + Retry-After header support.
# Increments the counter passed by [ref] so the caller can log the real retry count.
function Invoke-GraphWithRetry {
    param(
        [Parameter(Mandatory)][scriptblock]$ScriptBlock,
        [int]$MaxAttempts = 5,
        [int]$InitialDelaySeconds = 2,
        [ref]$AttemptCounter
    )
    $attempt = 0
    $delay   = $InitialDelaySeconds
    while ($true) {
        $attempt++
        if ($PSBoundParameters.ContainsKey('AttemptCounter')) { $AttemptCounter.Value = $attempt }
        try {
            return & $ScriptBlock
        }
        catch {
            $ex         = $_.Exception
            $status     = $null
            $retryAfter = $null

            if ($ex.PSObject.Properties.Name -contains 'Response' -and $ex.Response) {
                try { $status = [int]$ex.Response.StatusCode } catch { }
                try {
                    $ra = $ex.Response.Headers['Retry-After']
                    if ($ra) { $retryAfter = [int]$ra }
                } catch { }
            }

            $isRetryable = ($status -in 429, 500, 502, 503, 504) -or
                           ($ex.Message -match 'throttl|timeout|temporarily|service unavailable')

            if (-not $isRetryable -or $attempt -ge $MaxAttempts) { throw }

            $wait = if ($retryAfter) { $retryAfter } else { $delay }
            Write-Host "    [RETRY $attempt/$MaxAttempts] HTTP $status - sleeping ${wait}s..." -ForegroundColor DarkYellow
            Start-Sleep -Seconds $wait
            $delay = [Math]::Min($delay * 2, 60)
        }
    }
}

# Does the issued token actually WORK? Connect-MgGraph can report success and still hand
# back a credential that throws on first use (see the -UseDeviceCode / Azure.Identity WAM
# regression in Graph SDK 2.34+). One cheap read settles it before we touch any data.
function Test-GraphToken {
    try {
        $null = Get-MgUser -Top 1 -Property 'Id' -ErrorAction Stop
        return [PSCustomObject]@{ Ok = $true; Message = '' }
    }
    catch {
        return [PSCustomObject]@{ Ok = $false; Message = [string]$_.Exception.Message }
    }
}

# Is this failure the auth layer rather than the data? Used to tell "this user does not
# exist" (skip the row) apart from "our token is dead" (abort the whole run).
function Test-IsAuthFailure {
    param([AllowEmptyString()][string]$Message)
    if ([string]::IsNullOrEmpty($Message)) { return $false }
    return ($Message -match 'authentication failed|Object reference not set to an instance|InvalidAuthenticationToken|Access token has expired|CompactToken|401 \(Unauthorized\)|Lifetime validation failed|AADSTS')
}

function Get-FileHashSHA256 {
    param([string]$Path)
    try { return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash } catch { return 'HASH_ERROR' }
}

function Write-PhaseHeader {
    param([string]$Title)
    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host " $Title" -ForegroundColor Cyan
    Write-Host "============================================================" -ForegroundColor Cyan
}

# Ensure a module is present + imported. Prompts (Y/N) before installing.
function Ensure-Module {
    param(
        [Parameter(Mandatory)][string]$Name,
        [switch]$AutoInstall
    )

    if (Get-Module -ListAvailable -Name $Name) {
        try {
            Import-Module $Name -ErrorAction Stop
            return $true
        }
        catch {
            Write-Host "[FAIL] '$Name' is installed but failed to import: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "This often indicates Graph SDK version drift. Recommended clean reinstall:" -ForegroundColor Yellow
            Write-Host "  Uninstall-Module Microsoft.Graph -AllVersions -Force" -ForegroundColor Yellow
            Write-Host "  Get-InstalledModule Microsoft.Graph* -AllVersions | Uninstall-Module -Force" -ForegroundColor Yellow
            Write-Host "  Install-Module Microsoft.Graph -Scope CurrentUser -Force" -ForegroundColor Yellow
            return $false
        }
    }

    Write-Host ""
    Write-Host "[MISSING] Module '$Name' is not installed." -ForegroundColor Yellow
    $doInstall = $AutoInstall.IsPresent
    if (-not $doInstall) {
        $ans = Read-Host "Install '$Name' now for CurrentUser? (Y/N)"
        $doInstall = ($ans -match '^(y|yes)$')
    }
    if (-not $doInstall) {
        Write-Host "[FAIL] Install declined. Cannot continue without '$Name'." -ForegroundColor Red
        return $false
    }

    try {
        Write-Host "Installing '$Name' (Scope: CurrentUser)..." -ForegroundColor Cyan
        Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Import-Module $Name -ErrorAction Stop
        $installed = Get-Module $Name | Select-Object -First 1
        Write-Host "  [OK] Installed and imported: $Name v$($installed.Version)" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "[FAIL] Install of '$Name' failed: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Verify internet access, PSGallery trust, and permissions." -ForegroundColor Yellow
        return $false
    }
}

# Append-only flush of the audit log. Returns the new flushed-row count.
#
# v3.4.0 and earlier rewrote the ENTIRE log every 10 rows. At 23,000 rows that is O(n^2)
# file I/O - by the end each flush rewrites 23,000 rows, ~26 million row-writes over a run.
# Appending only the new slice makes flushing O(1) per row regardless of batch size.
function Save-LogIncremental {
    param(
        [System.Collections.Generic.List[object]]$Log,
        [Parameter(Mandatory)][string]$Path,
        [int]$AlreadyFlushed = 0
    )
    if ($null -eq $Log -or $Log.Count -le $AlreadyFlushed) { return $AlreadyFlushed }
    try {
        $new = $Log.GetRange($AlreadyFlushed, $Log.Count - $AlreadyFlushed)
        if ($AlreadyFlushed -eq 0) {
            $new | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8 -Force -ErrorAction Stop
        } else {
            $new | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8 -Append -ErrorAction Stop
        }
        return $Log.Count
    }
    catch { return $AlreadyFlushed }
}

# Human-readable duration for ETAs.
function Format-Duration {
    param([double]$Seconds)
    if ($Seconds -lt 0) { $Seconds = 0 }
    $ts = [TimeSpan]::FromSeconds([Math]::Round($Seconds))
    if ($ts.TotalHours -ge 1) { return ('{0}h {1:00}m {2:00}s' -f [int]$ts.TotalHours, $ts.Minutes, $ts.Seconds) }
    if ($ts.TotalMinutes -ge 1) { return ('{0}m {1:00}s' -f [int]$ts.TotalMinutes, $ts.Seconds) }
    return ('{0}s' -f $ts.Seconds)
}

# ---------- JSON batching ----------
# Graph's $batch endpoint takes up to 20 sub-requests per call and runs them in PARALLEL
# server-side. Each sub-request still counts individually against throttling, so this does
# not raise the ceiling (Entra allows ~3,000 writes / 2.5 min per app per tenant = ~20/sec);
# it removes the per-request round-trip wait that was capping us at 2-3 writes/sec.
$GraphBatchMax        = 20     # hard Graph limit
$TargetWritesPerSec   = 18     # ~90% of the 20/sec Entra write ceiling, leaves headroom

# Pull '[code] message' out of a batch sub-response body.
function Get-BatchErrorText {
    param($Body)
    if ($null -eq $Body) { return 'Unknown error (empty body)' }
    try {
        if ($Body -is [System.Collections.IDictionary] -and $Body.Contains('error')) {
            $e    = $Body['error']
            $code = [string]$e['code']
            $msg  = [string]$e['message']
            if ($code -or $msg) { return "[$code] $msg" }
        }
    } catch { }
    return ([string]$Body)
}

# Send one JSON batch and return a per-sub-request result map.
# Retries ONLY the throttled/5xx sub-requests, honouring the per-sub-response Retry-After.
function Invoke-GraphBatch {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][array]$Requests,
        [int]$MaxAttempts = 5
    )
    $results   = @{}
    $pending   = @($Requests)
    $attempt   = 0
    $throttled = $false

    while ($pending.Count -gt 0 -and $attempt -lt $MaxAttempts) {
        $attempt++
        $payload = @{ requests = $pending } | ConvertTo-Json -Depth 8 -Compress

        $resp = Invoke-MgGraphRequest -Method POST -Uri 'https://graph.microsoft.com/v1.0/$batch' `
                    -Body $payload -ContentType 'application/json' -OutputType HashTable -ErrorAction Stop

        $retryThese    = @()
        $maxRetryAfter = 0

        foreach ($r in $resp['responses']) {
            $id     = [string]$r['id']
            $status = 0
            try { $status = [int]$r['status'] } catch { }

            $bodyObj = $null
            if ($r.Contains('body')) { $bodyObj = $r['body'] }

            $ra = 0
            try {
                if ($r.Contains('headers') -and $r['headers'] -and $r['headers']['Retry-After']) {
                    $ra = [int]$r['headers']['Retry-After']
                }
            } catch { }

            if ($status -in 429, 500, 502, 503, 504) {
                $throttled = $true
                if ($ra -gt $maxRetryAfter) { $maxRetryAfter = $ra }
                $retryThese += @($pending | Where-Object { [string]$_.id -eq $id })
                continue
            }

            $results[$id] = [PSCustomObject]@{ Status = $status; Body = $bodyObj }
        }

        $pending = @($retryThese)
        if ($pending.Count -gt 0) {
            $wait = if ($maxRetryAfter -gt 0) { $maxRetryAfter } else { [int][Math]::Min(30, [Math]::Pow(2, $attempt)) }
            Write-Host ("    [BATCH RETRY {0}/{1}] {2} sub-request(s) throttled - waiting {3}s" -f `
                $attempt, $MaxAttempts, $pending.Count, $wait) -ForegroundColor DarkYellow
            Start-Sleep -Seconds $wait
        }
    }

    # Give up honestly rather than silently dropping rows
    foreach ($p in $pending) {
        $results[[string]$p.id] = [PSCustomObject]@{
            Status = 429
            Body   = @{ error = @{ code = 'throttledRetryExhausted'
                                   message = "Still throttled after $MaxAttempts batch attempts" } }
        }
    }

    return [PSCustomObject]@{ Results = $results; Throttled = $throttled; Attempts = $attempt }
}

# One audit row. Shared by the sequential and batch write paths so the schema cannot drift.
function New-AuditRow {
    param(
        [Parameter(Mandatory)]$Row,
        [Parameter(Mandatory)][string]$Status,
        [AllowEmptyString()][string]$Message = '',
        [AllowEmptyString()][string]$HttpStatusCode = '',
        [AllowEmptyString()][string]$GraphErrorCode = '',
        [AllowEmptyString()][string]$GraphRequestId = '',
        [int]$AttemptCount = 1,
        [int]$DurationMs = 0,
        [AllowEmptyString()][string]$RevokeStatus = '',
        [AllowEmptyString()][string]$RevokeMessage = '',
        [Parameter(Mandatory)][string]$CorrelationId,
        [Parameter(Mandatory)][string]$Operator,
        [Parameter(Mandatory)][string]$Tenant,
        [Parameter(Mandatory)][string]$Version,
        [AllowEmptyString()][string]$GenStyle = '',
        [AllowEmptyString()][string]$WriteMode = ''
    )
    return [PSCustomObject]@{
        TimestampLocal        = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        TimestampUtc          = (Get-Date).ToUniversalTime().ToString('o')
        CorrelationId         = $CorrelationId
        Row                   = $Row.Row
        UpnRaw                = $Row.UpnRaw
        UpnCleaned            = $Row.UpnClean
        SanitizationApplied   = if ($Row.UpnChanged) { 'Yes' } else { 'No' }
        UserObjectId          = $Row.UserObjectId
        DisplayName           = $Row.DisplayName
        PasswordSource        = $Row.PasswordSource
        PasswordLength        = $Row.Password.Length
        PasswordFingerprint   = (Get-SecretFingerprint -Secret $Row.Password -Salt $CorrelationId)
        PasswordGenStyle      = if ($Row.PasswordSource -eq 'Generated') { $GenStyle } else { '' }
        ForceChangeNextSignIn = if ($Status -eq 'Success') { 'Yes' } else { '' }
        SessionRevokeStatus   = $RevokeStatus
        SessionRevokeMessage  = $RevokeMessage
        Status                = $Status
        Message               = $Message
        HttpStatusCode        = $HttpStatusCode
        GraphErrorCode        = $GraphErrorCode
        GraphRequestId        = $GraphRequestId
        AttemptCount          = $AttemptCount
        DurationMs            = $DurationMs
        WriteMode             = $WriteMode
        OperatorAccount       = $Operator
        TenantId              = $Tenant
        ScriptVersion         = $Version
    }
}

# Page the whole directory once and index it by UPN.
#
# Pre-flight needs six properties per target user. Asking Graph one user at a time costs one
# round trip per row - fine for 30 rows, ruinous for 23,000. A paged read at 999/page turns
# 23,145 requests into ~24, and every lookup afterwards is an in-memory hash hit.
function Get-TenantUserIndex {
    param([int]$PageSize = 999)

    $index  = [System.Collections.Generic.Dictionary[string,object]]::new([StringComparer]::OrdinalIgnoreCase)
    $select = 'id,userPrincipalName,accountEnabled,userType,onPremisesSyncEnabled,displayName'
    $uri    = "/v1.0/users?`$select=$select&`$top=$PageSize"
    $page   = 0
    $start  = Get-Date

    while ($uri) {
        $page++
        $resp = Invoke-GraphWithRetry -ScriptBlock {
            Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType HashTable -ErrorAction Stop
        }

        foreach ($u in $resp['value']) {
            $upn = [string]$u['userPrincipalName']
            if (-not $upn) { continue }
            $index[$upn] = [PSCustomObject]@{
                Id                    = [string]$u['id']
                UserPrincipalName     = $upn
                AccountEnabled        = $u['accountEnabled']
                UserType              = [string]$u['userType']
                OnPremisesSyncEnabled = $u['onPremisesSyncEnabled']
                DisplayName           = [string]$u['displayName']
            }
        }

        $secs = ((Get-Date) - $start).TotalSeconds
        Write-Host ("    page {0}: {1} users indexed ({2})" -f $page, $index.Count, (Format-Duration $secs)) -ForegroundColor DarkGray

        $uri = if ($resp.ContainsKey('@odata.nextLink')) { [string]$resp['@odata.nextLink'] } else { $null }
    }

    return $index
}

# Best-effort: lock the credentials file down to the current user (Windows only).
function Set-SecretFileAcl {
    param([Parameter(Mandatory)][string]$Path)
    if (-not $IsWindows) { return $false }
    try {
        $acl = Get-Acl -LiteralPath $Path
        $acl.SetAccessRuleProtection($true, $false)   # break inheritance, drop inherited rules
        foreach ($r in @($acl.Access)) { [void]$acl.RemoveAccessRule($r) }
        $me   = "$env:USERDOMAIN\$env:USERNAME"
        $rule = New-Object System.Security.AccessControl.FileSystemAccessRule($me, 'FullControl', 'Allow')
        $acl.SetAccessRule($rule)
        Set-Acl -LiteralPath $Path -AclObject $acl -ErrorAction Stop
        return $true
    }
    catch { return $false }
}


# ============================================================
# PHASE 1 - Configure Paths (Input CSV + Log Folder + Credentials Folder)
# ============================================================
Write-PhaseHeader "PHASE 1: Configure Paths"

Write-Host ""
Write-Host "Correlation ID : $CorrelationId" -ForegroundColor Magenta
Write-Host "Script version : $ScriptVersion" -ForegroundColor DarkGray

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

$DefaultCsvPath   = Join-Path $ScriptDir 'Bulk-ResetPasswords-Input.csv'
$DefaultLogFolder = Join-Path $ScriptDir 'Logs'

Write-Host ""
Write-Host "Default CSV path:" -ForegroundColor Gray
Write-Host "  $DefaultCsvPath" -ForegroundColor DarkGray
$InputCsv = Read-Host "Enter CSV file path (or press Enter to use default)"
$CsvPath  = if ([string]::IsNullOrWhiteSpace($InputCsv)) { $DefaultCsvPath } else { Convert-PathSafe $InputCsv }

Write-Host ""
Write-Host "Default log folder:" -ForegroundColor Gray
Write-Host "  $DefaultLogFolder" -ForegroundColor DarkGray
$InputLogFolder = Read-Host "Enter log output folder (or press Enter to use default)"
$LogFolder      = if ([string]::IsNullOrWhiteSpace($InputLogFolder)) { $DefaultLogFolder } else { Convert-PathSafe $InputLogFolder }

if (-not (Test-Path -LiteralPath $CsvPath)) {
    Write-Host ""
    Write-Host "[FAIL] CSV file not found at:" -ForegroundColor Red
    Write-Host "  $CsvPath" -ForegroundColor Red
    $codes = ([char[]]$CsvPath | Select-Object -First 5 | ForEach-Object { [int]$_ }) -join ','
    Write-Host "  First 5 char codes: $codes  (normal 'C:\' starts with 67,58,92)" -ForegroundColor DarkGray
    return
}

if (-not (Test-Path -LiteralPath $LogFolder)) {
    try {
        New-Item -Path $LogFolder -ItemType Directory -Force | Out-Null
        Write-Host "Created log folder: $LogFolder" -ForegroundColor DarkGray
    }
    catch {
        Write-Host "[FAIL] Could not create log folder '$LogFolder': $($_.Exception.Message)" -ForegroundColor Red
        return
    }
}

# --- Credentials folder: separate destination for the cleartext password file ---
Write-Host ""
Write-Host "The CREDENTIALS file contains cleartext temporary passwords." -ForegroundColor Yellow
if ($LogFolder -match 'OneDrive|SharePoint|Dropbox|Google Drive|Box') {
    Write-Host "[WARN] Your log folder looks CLOUD-SYNCED:" -ForegroundColor Yellow
    Write-Host "       $LogFolder" -ForegroundColor DarkYellow
    Write-Host "       Cleartext passwords written there will replicate to the cloud and to" -ForegroundColor DarkYellow
    Write-Host "       version history. A local folder (e.g. C:\Temp\Creds) is strongly preferred." -ForegroundColor DarkYellow
}
$InputCredFolder = Read-Host "Enter folder for the CREDENTIALS file (or press Enter to reuse the log folder)"
$CredFolder      = if ([string]::IsNullOrWhiteSpace($InputCredFolder)) { $LogFolder } else { Convert-PathSafe $InputCredFolder }

if (-not (Test-Path -LiteralPath $CredFolder)) {
    try {
        New-Item -Path $CredFolder -ItemType Directory -Force | Out-Null
        Write-Host "Created credentials folder: $CredFolder" -ForegroundColor DarkGray
    }
    catch {
        Write-Host "[FAIL] Could not create credentials folder '$CredFolder': $($_.Exception.Message)" -ForegroundColor Red
        return
    }
}

# Build output paths + fingerprint the input CSV
$TimeStamp       = Get-Date -Format "yyyyMMdd_HHmmss"
$LogPath         = Join-Path $LogFolder  "ResetPasswords_Log_$TimeStamp.csv"
$ManifestPath    = Join-Path $LogFolder  "ResetPasswords_RunSummary_$TimeStamp.json"
$CredentialsPath = Join-Path $CredFolder "ResetPasswords_CREDENTIALS_$TimeStamp.csv"
$FailedLogPath   = Join-Path $CredFolder "ResetPasswords_FAILED_RERUN_$TimeStamp.csv"
$CsvHash         = Get-FileHashSHA256 -Path $CsvPath

Write-Host ""
Write-Host "Paths configured:" -ForegroundColor Green
Write-Host "  CSV Input        : $CsvPath" -ForegroundColor Green
Write-Host "  CSV SHA256       : $CsvHash" -ForegroundColor DarkGray
Write-Host "  Audit log        : $LogPath" -ForegroundColor Green
Write-Host "     (no cleartext passwords - safe to share)" -ForegroundColor DarkGray
Write-Host "  Run manifest     : $ManifestPath" -ForegroundColor Green
Write-Host "  CREDENTIALS file : $CredentialsPath" -ForegroundColor Yellow
Write-Host "  Failed re-run    : $FailedLogPath" -ForegroundColor Yellow
Write-Host "     (both contain SECRETS - distribute, then delete)" -ForegroundColor DarkYellow


# ============================================================
# PHASE 2 - Connect to Microsoft Graph (Entra ID)
# ============================================================
Write-PhaseHeader "PHASE 2: Connect to Microsoft Graph"

$RequiredModules = @(
    'Microsoft.Graph.Users',
    'Microsoft.Graph.Users.Actions',              # Revoke-MgUserSignInSession lives here
    'Microsoft.Graph.Identity.DirectoryManagement'
)
Write-Host ""
Write-Host "Checking required modules..." -ForegroundColor Cyan
foreach ($Module in $RequiredModules) {
    if (-not (Ensure-Module -Name $Module)) { return }
}
Write-Host "[OK] All required modules ready." -ForegroundColor Green

# NO hardcoded default - always prompt
Write-Host ""
$InputTenantId = Read-Host "Enter Tenant ID (GUID)"
$TenantId = (Remove-InvisibleChars $InputTenantId).Trim()
if ([string]::IsNullOrWhiteSpace($TenantId)) {
    Write-Host "[FAIL] Tenant ID is required." -ForegroundColor Red
    return
}
if ($TenantId -notmatch '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$') {
    Write-Host "[FAIL] Tenant ID does not look like a valid GUID." -ForegroundColor Red
    return
}

# IMPORTANT: User.ReadWrite.All alone is NOT sufficient to write passwordProfile.
# Microsoft carved password reset out into the dedicated User-PasswordProfile.ReadWrite.All
# permission (Graph changelog, 2024-12-23). Without it every PATCH returns
# 403 Authorization_RequestDenied even for a Global Administrator.
$RequiredScopes = @(
    'User-PasswordProfile.ReadWrite.All',
    'User.RevokeSessions.All',
    'User.ReadWrite.All',
    'Directory.Read.All',
    'RoleManagement.Read.Directory'
)
$CriticalScope = 'User-PasswordProfile.ReadWrite.All'
# revokeSignInSessions accepts User.RevokeSessions.All (least privilege) OR the
# higher-privileged User.ReadWrite.All / Directory.ReadWrite.All.
$RevokeScopes  = @('User.RevokeSessions.All','User.ReadWrite.All','Directory.ReadWrite.All')

Write-Host ""
Write-Host "Scopes to be requested:" -ForegroundColor Cyan
Write-Host "  User-PasswordProfile.ReadWrite.All - REQUIRED to write passwordProfile (reset)" -ForegroundColor DarkGray
Write-Host "  User.RevokeSessions.All            - force sign-out (revokeSignInSessions), optional" -ForegroundColor DarkGray
Write-Host "  User.ReadWrite.All                 - read/write the user object" -ForegroundColor DarkGray
Write-Host "  Directory.Read.All                 - read UserType / OnPremisesSyncEnabled / AccountEnabled" -ForegroundColor DarkGray
Write-Host "  RoleManagement.Read.Directory      - detect privileged role holders before resetting" -ForegroundColor DarkGray

Write-Host ""
Write-Host "Sign-in mode:" -ForegroundColor Cyan
Write-Host "  [1] Interactive popup window  (RECOMMENDED - default)" -ForegroundColor White
Write-Host "  [2] Device code               (only for headless/SSH - see warning below)" -ForegroundColor White
Write-Host "  [3] Reuse existing session if already connected to this tenant" -ForegroundColor White
Write-Host ""
Write-Host "  [WARN] Graph SDK 2.34+ bumped Azure.Identity to use the WAM broker, which broke" -ForegroundColor DarkYellow
Write-Host "         device-code token caching. -UseDeviceCode signs in fine, then EVERY cmdlet" -ForegroundColor DarkYellow
Write-Host "         fails with 'DeviceCodeCredential authentication failed: Object reference" -ForegroundColor DarkYellow
Write-Host "         not set to an instance of an object.' (SDK issue #3495). Prefer [1]." -ForegroundColor DarkYellow
Write-Host ""
$SignInChoice = (Read-Host "Enter choice (1, 2, or 3 - press Enter for 1)").Trim()
if ([string]::IsNullOrWhiteSpace($SignInChoice)) { $SignInChoice = '1' }

$AlreadyConnected = $false
$Context          = $null

if ($SignInChoice -eq '3') {
    try {
        $ExistingContext = Get-MgContext -ErrorAction Stop
        if ($ExistingContext -and $ExistingContext.TenantId -eq $TenantId) {
            $AlreadyConnected = $true
            $Context = $ExistingContext
            Write-Host ""
            Write-Host "[OK] Reusing existing session." -ForegroundColor Green
            Write-Host "  Account : $($Context.Account)" -ForegroundColor Green
            Write-Host "  Tenant  : $($Context.TenantId)" -ForegroundColor Green
        } else {
            Write-Host "[WARN] No matching existing session - falling back to browser sign-in." -ForegroundColor DarkYellow
            $SignInChoice = '1'
        }
    }
    catch {
        Write-Host "[WARN] No existing session - falling back to browser sign-in." -ForegroundColor DarkYellow
        $SignInChoice = '1'
    }
}

if (-not $AlreadyConnected) {
    try {
        Write-Host ""
        Write-Host "Connecting to Microsoft Graph (Tenant: $TenantId)..." -ForegroundColor Cyan
        if ($SignInChoice -eq '2') {
            Write-Host "A device code will be shown. Open the URL in your browser and enter the code." -ForegroundColor DarkGray
            Connect-MgGraph -TenantId $TenantId -Scopes $RequiredScopes -UseDeviceCode -NoWelcome -ErrorAction Stop
        } else {
            Write-Host "An interactive popup window will appear for sign-in." -ForegroundColor DarkGray
            Connect-MgGraph -TenantId $TenantId -Scopes $RequiredScopes -NoWelcome -ErrorAction Stop
        }
        $Context = Get-MgContext
        Write-Host ""
        Write-Host "[OK] Connected successfully." -ForegroundColor Green
        Write-Host "  Account : $($Context.Account)" -ForegroundColor Green
        Write-Host "  Tenant  : $($Context.TenantId)" -ForegroundColor Green
    }
    catch {
        Write-Host "[FAIL] Failed to connect: $($_.Exception.Message)" -ForegroundColor Red
        return
    }
}

if (-not $Context -or [string]::IsNullOrWhiteSpace($Context.Account)) {
    Write-Host "[FAIL] Graph context is empty after connect. Aborting." -ForegroundColor Red
    return
}
$OperatorAccount = $Context.Account

# ---------- 2a. Smoke-test the token with a real Graph call ----------
# "Connected successfully" only means a token was issued. Under the SDK 2.34+ device-code
# regression the credential object is unusable and the first real call throws. Find out now,
# with one cheap read, instead of mislabelling all 26 users as "not found in tenant".
Write-Host ""
Write-Host "Smoke-testing the token with a live Graph call..." -ForegroundColor Cyan
$Smoke = Test-GraphToken

if (-not $Smoke.Ok -and (Test-IsAuthFailure -Message $Smoke.Message)) {
    Write-Host "  [FAIL] The token was issued but does not work:" -ForegroundColor Red
    Write-Host "         $($Smoke.Message)" -ForegroundColor DarkRed

    $IsDeviceCodeBug = ($Smoke.Message -match 'DeviceCodeCredential|Object reference not set to an instance')
    if ($IsDeviceCodeBug) {
        Write-Host ""
        Write-Host "  Diagnosis: known Microsoft Graph SDK regression (issue #3495)." -ForegroundColor Yellow
        Write-Host "  SDK 2.34+ bumped Azure.Identity to use the WAM broker, changing how tokens are" -ForegroundColor Yellow
        Write-Host "  cached. Device-code sign-in reports success, then every cmdlet fails like this." -ForegroundColor Yellow
    }

    # Auto-recover: the documented workaround is simply not to use device code.
    if ($SignInChoice -eq '2') {
        Write-Host ""
        Write-Host "  The interactive popup uses a different token path and is unaffected." -ForegroundColor Cyan
        $reAns = (Read-Host "  Reconnect now using the interactive popup instead? (Y/N)").Trim()
        if ($reAns -match '^(y|yes)$') {
            try {
                Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
                Write-Host "  Reconnecting interactively..." -ForegroundColor Cyan
                Connect-MgGraph -TenantId $TenantId -Scopes $RequiredScopes -NoWelcome -ErrorAction Stop
                $Context         = Get-MgContext
                $OperatorAccount = $Context.Account
                $SignInChoice    = '1'
                $Smoke           = Test-GraphToken
                if ($Smoke.Ok) {
                    Write-Host "  [OK] Reconnected interactively and the token works." -ForegroundColor Green
                    Write-Host "       Account: $OperatorAccount" -ForegroundColor Green
                }
            }
            catch {
                Write-Host "  [FAIL] Interactive reconnect failed: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }

    if (-not $Smoke.Ok) {
        Write-Host ""
        Write-Host "Cannot continue with an unusable token. Fix options, best first:" -ForegroundColor Cyan
        Write-Host "  1. Re-run this script and choose sign-in mode [1] Interactive popup." -ForegroundColor White
        Write-Host "  2. If you MUST use device code (headless/SSH), pin the SDK below the" -ForegroundColor White
        Write-Host "     regression, or try the newest build in case the fix has shipped:" -ForegroundColor White
        Write-Host "       Get-InstalledModule Microsoft.Graph* -AllVersions | Uninstall-Module -Force" -ForegroundColor DarkGray
        Write-Host "       Install-Module Microsoft.Graph -RequiredVersion 2.33.0 -Scope CurrentUser -Force" -ForegroundColor DarkGray
        Write-Host "  3. Use app-only certificate auth, which bypasses the broker entirely:" -ForegroundColor White
        Write-Host "       Connect-MgGraph -ClientId <appId> -TenantId <tenantId> -CertificateThumbprint <thumb>" -ForegroundColor DarkGray
        Write-Host "  4. Check for module version drift - mixed Microsoft.Graph.* versions, or Az /" -ForegroundColor White
        Write-Host "     PnP.PowerShell loaded first, cause the same class of failure:" -ForegroundColor White
        Write-Host "       Get-Module Microsoft.Graph* -ListAvailable | Select-Object Name,Version | Sort-Object Name" -ForegroundColor DarkGray
        Write-Host ""
        Write-Host "Aborting - nothing was read, nothing was changed, no files written." -ForegroundColor Red
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        return
    }
}
elseif (-not $Smoke.Ok) {
    # Not an auth signature - could be a missing read scope. Warn but let scope checks speak.
    Write-Host "  [WARN] Test read failed (not an auth error): $($Smoke.Message)" -ForegroundColor DarkYellow
}
else {
    Write-Host "  [OK] Token verified against Microsoft Graph." -ForegroundColor Green
}

# ---------- 2b. Verify the token ACTUALLY carries the scopes we asked for ----------
# Connect-MgGraph succeeds even when consent was only partially granted, so a request
# can look fine here and then 403 on every single row in Phase 5. Check up front.
Write-Host ""
Write-Host "Verifying granted scopes on the issued token..." -ForegroundColor Cyan
$GrantedScopes = @()
try { $GrantedScopes = @($Context.Scopes) } catch { }
$RevokeCapable = $true   # assumed until proven otherwise

if ($GrantedScopes.Count -eq 0) {
    Write-Host "  [WARN] Could not read granted scopes from the context - cannot pre-verify." -ForegroundColor DarkYellow
}
else {
    $MissingScopes = @($RequiredScopes | Where-Object { $_ -notin $GrantedScopes })

    # Force sign-out needs ANY ONE of the revoke-capable scopes, not specifically the least-privileged one
    $RevokeGranted = @($RevokeScopes | Where-Object { $_ -in $GrantedScopes })
    $RevokeCapable = ($RevokeGranted.Count -gt 0)
    if ($RevokeCapable) {
        # Don't nag about the dedicated scope when a higher-privileged one already covers it
        $MissingScopes = @($MissingScopes | Where-Object { $_ -ne 'User.RevokeSessions.All' })
    }
    foreach ($s in $RequiredScopes) {
        if ($s -in $GrantedScopes) { Write-Host "  [OK]      $s" -ForegroundColor Green }
        else                       { Write-Host "  [MISSING] $s" -ForegroundColor Red }
    }

    if ($CriticalScope -notin $GrantedScopes) {
        Write-Host ""
        Write-Host "[FAIL] The token does NOT carry '$CriticalScope'." -ForegroundColor Red
        Write-Host "       Every password reset would return 403 Authorization_RequestDenied." -ForegroundColor Red
        Write-Host ""
        Write-Host "Why: Microsoft split password reset out of User.ReadWrite.All into a dedicated" -ForegroundColor Yellow
        Write-Host "     permission (Graph changelog 2024-12-23). User.ReadWrite.All is no longer enough." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "How to fix - run this once, accept the consent prompt, then re-run this script:" -ForegroundColor Cyan
        Write-Host "     Disconnect-MgGraph" -ForegroundColor White
        Write-Host "     Connect-MgGraph -TenantId $TenantId ``" -ForegroundColor White
        Write-Host "       -Scopes '$($RequiredScopes -join "','")'" -ForegroundColor White
        Write-Host ""
        Write-Host "If the consent prompt does not appear, a Global Administrator must grant admin" -ForegroundColor Yellow
        Write-Host "consent for '$CriticalScope' to the Microsoft Graph Command Line Tools app:" -ForegroundColor Yellow
        Write-Host "     Entra admin center > Applications > Enterprise applications >" -ForegroundColor DarkGray
        Write-Host "     Microsoft Graph Command Line Tools > Permissions > Grant admin consent" -ForegroundColor DarkGray
        Write-Host ""
        Write-Host "Aborting now - nothing was changed and no credentials file was written." -ForegroundColor Red
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        return
    }
    if ($MissingScopes.Count -gt 0) {
        Write-Host ""
        Write-Host "[WARN] Non-critical scope(s) missing: $($MissingScopes -join ', ')" -ForegroundColor DarkYellow
        Write-Host "       Pre-flight enrichment or admin detection may be degraded." -ForegroundColor DarkYellow
    }
    if ($RevokeCapable) {
        Write-Host "  [OK] Force sign-out available via: $($RevokeGranted -join ', ')" -ForegroundColor Green
    } else {
        Write-Host "  [INFO] Force sign-out NOT available - none of these scopes granted:" -ForegroundColor DarkYellow
        Write-Host "         $($RevokeScopes -join ', ')" -ForegroundColor DarkGray
        Write-Host "         The force sign-out option will be disabled for this run." -ForegroundColor DarkYellow
    }
}

# ---------- 2c. Advisory: does the operator hold a reset-capable directory role? ----------
# A valid scope is necessary but not sufficient - Entra also requires a suitable role,
# and a PIM-eligible role that has not been ACTIVATED behaves exactly like no role at all.
$ResetCapableRoles = @(
    'Global Administrator',
    'Privileged Authentication Administrator',
    'User Administrator',
    'Helpdesk Administrator',
    'Password Administrator',
    'Authentication Administrator'
)
Write-Host ""
Write-Host "Checking your active directory roles..." -ForegroundColor Cyan
$roleCheckStart = Get-Date
try {
    $OperatorUser  = Get-MgUser -UserId $OperatorAccount -Property 'Id,UserPrincipalName' -ErrorAction Stop
    $OperatorRoles = @()

    # ONE call. v3.3.0 enumerated every activated directory role and then paged every member
    # of each one - dozens of requests with no output, which reads exactly like a hang.
    # memberOf returns the roles the operator is ACTIVELY in, which is precisely the question.
    $memberOf = @(Get-MgUserMemberOf -UserId $OperatorUser.Id -All -ErrorAction Stop)
    foreach ($m in $memberOf) {
        try {
            if ([string]$m.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.directoryRole') {
                $dn = [string]$m.AdditionalProperties['displayName']
                if ($dn) { $OperatorRoles += $dn }
            }
        } catch { }
    }
    $OperatorRoles = @($OperatorRoles | Sort-Object -Unique)

    Write-Host ("  (took {0:N1}s)" -f ((Get-Date) - $roleCheckStart).TotalSeconds) -ForegroundColor DarkGray
    if ($OperatorRoles.Count -eq 0) {
        Write-Host "[WARN] '$OperatorAccount' holds NO ACTIVE directory roles." -ForegroundColor DarkYellow
        Write-Host "       If your role is PIM-eligible you must ACTIVATE it first, then reconnect." -ForegroundColor DarkYellow
    } else {
        Write-Host "Active directory roles for ${OperatorAccount}:" -ForegroundColor Cyan
        $OperatorRoles | ForEach-Object { Write-Host "  - $_" -ForegroundColor DarkGray }
        $capable = @($OperatorRoles | Where-Object { $_ -in $ResetCapableRoles })
        if ($capable.Count -gt 0) {
            Write-Host "  [OK] Reset-capable role present: $($capable -join ', ')" -ForegroundColor Green
        } else {
            Write-Host "  [WARN] None of these roles can reset user passwords." -ForegroundColor DarkYellow
            Write-Host "         Need one of: $($ResetCapableRoles -join ', ')" -ForegroundColor DarkYellow
        }
        if ('Privileged Authentication Administrator' -notin $OperatorRoles -and
            'Global Administrator' -notin $OperatorRoles) {
            Write-Host "  [NOTE] Without Privileged Authentication Administrator or Global Administrator," -ForegroundColor DarkGray
            Write-Host "         resets against OTHER ADMIN accounts will be denied by design." -ForegroundColor DarkGray
        }
    }
} catch {
    Write-Host ""
    Write-Host "[WARN] Could not evaluate operator roles: $($_.Exception.Message)" -ForegroundColor DarkYellow
    Write-Host "       Continuing - this check is advisory only." -ForegroundColor DarkYellow
}


# ============================================================
# PHASE 3 - Verify CSV Contents + Sanitize + Validate
# ============================================================
Write-PhaseHeader "PHASE 3: Verify CSV Contents"

try {
    $Users = @(Import-Csv -LiteralPath $CsvPath -Encoding UTF8 -ErrorAction Stop)
}
catch {
    Write-Host "[FAIL] Failed to read CSV: $($_.Exception.Message)" -ForegroundColor Red
    return
}

if ($Users.Count -eq 0) {
    Write-Host "[FAIL] CSV file is empty: $CsvPath" -ForegroundColor Red
    return
}

# 'upn' is mandatory. 'password' is now OPTIONAL - blank/absent means "generate".
$ActualColumns     = @($Users[0].PSObject.Properties.Name)
if ('upn' -notin $ActualColumns) {
    Write-Host "[FAIL] CSV is missing the required column: upn" -ForegroundColor Red
    Write-Host "Expected headers: upn  (optional: password, displayname)" -ForegroundColor Yellow
    Write-Host "Found headers   : $($ActualColumns -join ', ')" -ForegroundColor Yellow
    return
}
$HasPasswordColumn = ('password' -in $ActualColumns)

Write-Host ""
if ($HasPasswordColumn) {
    Write-Host "[OK] 'password' column present - CSV-only and Hybrid modes are available." -ForegroundColor Green
} else {
    Write-Host "[INFO] No 'password' column - passwords will be GENERATED for every row." -ForegroundColor Cyan
}

# Sanitize + validate every row
$ProcessedRows  = [System.Collections.Generic.List[object]]::new()
$SanitizedCount = 0
$RowNum         = 0

foreach ($u in $Users) {
    $RowNum++
    $rawUpn = if ($null -ne $u.upn) { [string]$u.upn } else { '' }
    $rawPwd = if ($HasPasswordColumn -and $null -ne $u.password) { ([string]$u.password).Trim() } else { '' }

    $upnResult = Convert-UpnToCanonical -Raw $rawUpn
    if ($upnResult.Changed) { $SanitizedCount++ }

    $ProcessedRows.Add([PSCustomObject]@{
        Row              = $RowNum
        UpnRaw           = $rawUpn
        UpnClean         = $upnResult.Clean
        UpnChanged       = $upnResult.Changed
        UpnValid         = $upnResult.Valid
        UpnReason        = $upnResult.Reason
        CsvPassword      = $rawPwd          # as supplied (may be empty)
        Password         = ''               # resolved in Phase 4
        PasswordSource   = ''               # 'CSV' | 'Generated'
        PasswordIssues   = ''
        UserObjectId     = ''
        DisplayName      = ''
        PreflightStatus  = 'Pending'
        PreflightMessage = ''
    })
}

# Duplicate UPN detection
$Duplicates = @($ProcessedRows | Where-Object { $_.UpnClean } |
                Group-Object UpnClean | Where-Object { $_.Count -gt 1 })

if ($SanitizedCount -gt 0) {
    Write-Host ""
    Write-Host "[WARN] $SanitizedCount UPN(s) required sanitization (invisible/whitespace/accented chars):" -ForegroundColor Yellow
    $changedRows = @($ProcessedRows | Where-Object { $_.UpnChanged })
    $changedRows | Select-Object -First 10 Row, UpnRaw, UpnClean | Format-Table -AutoSize | Out-Host
    if ($changedRows.Count -gt 10) {
        Write-Host "  ... and $($changedRows.Count - 10) more (see audit log for the full list)" -ForegroundColor DarkGray
    }
}

Write-Host ""
Write-Host "CSV loaded. Total rows: $($ProcessedRows.Count)" -ForegroundColor Green
Write-Host "Preview:" -ForegroundColor Cyan
Write-Host "------------------------------------------------------------" -ForegroundColor DarkGray
$ProcessedRows | Select-Object -First 20 Row,
    @{n='UPN';e={ if ($_.UpnClean) { $_.UpnClean } else { '<INVALID>' } }},
    @{n='CsvPwd';e={ if ($_.CsvPassword) { "yes ($($_.CsvPassword.Length) ch)" } else { '<blank>' } }} |
    Format-Table -AutoSize | Out-Host
if ($ProcessedRows.Count -gt 20) {
    Write-Host "  (showing first 20 of $($ProcessedRows.Count))" -ForegroundColor DarkGray
}
Write-Host "------------------------------------------------------------" -ForegroundColor DarkGray

# ---------- 3a. Row range: work a slice of the SAME file, no manual splitting ----------
# Row numbers are the ORIGINAL CSV line numbers and are preserved in every log, so slices
# from one file reconcile cleanly against each other.
$TotalCsvRows = $ProcessedRows.Count
Write-Host ""
Write-Host "Row range - process only part of this CSV?" -ForegroundColor Cyan
Write-Host "  The file has $TotalCsvRows data row(s). Row numbers below are original CSV rows," -ForegroundColor DarkGray
Write-Host "  so slices of one file stay reconcilable (e.g. 1-1000, then 1001-2000)." -ForegroundColor DarkGray
$rangeFrom = 1
$rangeTo   = $TotalCsvRows

$fromInput = (Read-Host "  Start at row (press Enter for 1)").Trim()
if (-not [string]::IsNullOrWhiteSpace($fromInput)) {
    $p = 0
    if ([int]::TryParse($fromInput, [ref]$p) -and $p -ge 1) { $rangeFrom = $p }
    else { Write-Host "  [WARN] Invalid start row - using 1." -ForegroundColor DarkYellow }
}
$toInput = (Read-Host "  End at row (press Enter for $TotalCsvRows = to the end)").Trim()
if (-not [string]::IsNullOrWhiteSpace($toInput)) {
    $p = 0
    if ([int]::TryParse($toInput, [ref]$p) -and $p -ge $rangeFrom) { $rangeTo = [Math]::Min($p, $TotalCsvRows) }
    else { Write-Host "  [WARN] Invalid end row - using $TotalCsvRows." -ForegroundColor DarkYellow }
}

if ($rangeFrom -gt 1 -or $rangeTo -lt $TotalCsvRows) {
    $ProcessedRows = [System.Collections.Generic.List[object]]::new(
        [object[]]@($ProcessedRows | Where-Object { $_.Row -ge $rangeFrom -and $_.Row -le $rangeTo }))
    Write-Host "  Selected rows $rangeFrom-$rangeTo : $($ProcessedRows.Count) row(s) of $TotalCsvRows." -ForegroundColor Yellow
    if ($ProcessedRows.Count -eq 0) {
        Write-Host "[FAIL] The selected range contains no rows." -ForegroundColor Red
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        return
    }
} else {
    Write-Host "  Processing all $TotalCsvRows row(s)." -ForegroundColor DarkGray
}

# ---------- 3b. Resume: skip users already reset successfully in a previous run ----------
# The point of a multi-hour job is that it must be restartable. Point this at the audit log
# from the interrupted run and every UPN already marked Success is skipped.
$ResumeSkipUpns = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$ResumeLogPath  = ''
Write-Host ""
Write-Host "Resume from a previous run?" -ForegroundColor Cyan
Write-Host "  Give the path to an earlier ResetPasswords_Log_*.csv and every user already marked" -ForegroundColor DarkGray
Write-Host "  'Success' in it will be skipped here. Leave blank for a fresh run." -ForegroundColor DarkGray
$resumeInput = Convert-PathSafe (Read-Host "  Previous audit log path (Enter to skip)")

if ($resumeInput) {
    if (-not (Test-Path -LiteralPath $resumeInput)) {
        Write-Host "  [WARN] Not found - continuing WITHOUT resume: $resumeInput" -ForegroundColor DarkYellow
    }
    else {
        try {
            $prev = @(Import-Csv -LiteralPath $resumeInput -ErrorAction Stop)
            $prevCols = if ($prev.Count -gt 0) { @($prev[0].PSObject.Properties.Name) } else { @() }
            if ('Status' -notin $prevCols -or 'UpnCleaned' -notin $prevCols) {
                Write-Host "  [WARN] That file has no Status/UpnCleaned columns - is it an audit log?" -ForegroundColor DarkYellow
                Write-Host "  [WARN] Continuing WITHOUT resume." -ForegroundColor DarkYellow
            }
            else {
                foreach ($p in $prev) {
                    if ([string]$p.Status -eq 'Success' -and $p.UpnCleaned) {
                        [void]$ResumeSkipUpns.Add([string]$p.UpnCleaned)
                    }
                }
                $ResumeLogPath = $resumeInput
                Write-Host "  Resume source: $ResumeLogPath" -ForegroundColor Green
                Write-Host "  $($ResumeSkipUpns.Count) previously successful user(s) will be skipped." -ForegroundColor Green
                $willSkip = @($ProcessedRows | Where-Object { $ResumeSkipUpns.Contains($_.UpnClean) }).Count
                Write-Host "  $willSkip of the $($ProcessedRows.Count) selected row(s) match and will be skipped." -ForegroundColor Yellow
            }
        }
        catch {
            Write-Host "  [WARN] Could not read that log: $($_.Exception.Message)" -ForegroundColor DarkYellow
            Write-Host "  [WARN] Continuing WITHOUT resume." -ForegroundColor DarkYellow
        }
    }
}

$BlankPwdCount = @($ProcessedRows | Where-Object { -not $_.CsvPassword }).Count
Write-Host "Rows with a CSV password : $($ProcessedRows.Count - $BlankPwdCount)" -ForegroundColor DarkGray
Write-Host "Rows with a blank password: $BlankPwdCount" -ForegroundColor DarkGray

if ($Duplicates.Count -gt 0) {
    Write-Host ""
    Write-Host "[WARN] $($Duplicates.Count) duplicate UPN(s) detected (first row wins, later rows skipped):" -ForegroundColor Yellow
    $Duplicates | ForEach-Object {
        Write-Host ("  {0}  (rows: {1})" -f $_.Name, (($_.Group.Row) -join ', ')) -ForegroundColor DarkYellow
    }
}


# ============================================================
# PHASE 4 - Password Source + Pre-flight + Confirmation
# ============================================================
Write-PhaseHeader "PHASE 4: Password Source, Pre-flight & Confirmation"

# ---------- 4a. Execution mode ----------
Write-Host ""
Write-Host "Choose execution mode:" -ForegroundColor Cyan
Write-Host "  [Y] Yes      - proceed with password resets" -ForegroundColor White
Write-Host "  [T] Test     - dry-run: validate UPNs + generate passwords, write NOTHING" -ForegroundColor White
Write-Host "  [N] No       - cancel and disconnect" -ForegroundColor White
Write-Host ""
$ModeChoice = (Read-Host "Enter choice (Y / T / N)").Trim().ToUpper()

if ($ModeChoice -eq 'N' -or [string]::IsNullOrWhiteSpace($ModeChoice)) {
    Write-Host ""
    Write-Host "Operation cancelled. No changes were made." -ForegroundColor Red
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    return
}
$IsDryRun = ($ModeChoice -eq 'T')

# ---------- 4b. Password source ----------
Write-Host ""
Write-Host "Password source:" -ForegroundColor Cyan
Write-Host "  [1] CSV only  - use the 'password' column for every row" -ForegroundColor White
Write-Host "  [2] Generate  - ignore the CSV password column, generate a unique password per user" -ForegroundColor White
Write-Host "  [3] Hybrid    - generate ONLY where the CSV password is blank/missing  (default)" -ForegroundColor White
Write-Host ""
$SourceChoice = (Read-Host "Enter choice (1, 2, or 3 - press Enter for 3)").Trim()
if ([string]::IsNullOrWhiteSpace($SourceChoice)) { $SourceChoice = '3' }
if ($SourceChoice -notin '1','2','3') {
    Write-Host "[FAIL] Invalid choice '$SourceChoice'." -ForegroundColor Red
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    return
}
if ($SourceChoice -eq '1' -and -not $HasPasswordColumn) {
    Write-Host "[FAIL] CSV-only mode selected but the CSV has no 'password' column." -ForegroundColor Red
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    return
}
$PasswordMode = switch ($SourceChoice) { '1' { 'CsvOnly' } '2' { 'GenerateAll' } '3' { 'Hybrid' } }

# ---------- 4c. Generation settings ----------
$GenStyle      = 'Readable'
$GenLength     = 14
$GenWordCount  = 3
$GenDigitCount = 4
$GenEntropy    = 0

if ($PasswordMode -ne 'CsvOnly') {
    Write-Host ""
    Write-Host "Password style:" -ForegroundColor Cyan
    Write-Host "  [1] Readable  - e.g. $(New-ReadablePassword -WordCount 3 -DigitCount 4)" -ForegroundColor White
    Write-Host "                  easy to dictate on the phone, print on a handout, type on a tablet" -ForegroundColor DarkGray
    Write-Host "  [2] Random    - e.g. $(New-RandomStringPassword -Length 14)" -ForegroundColor White
    Write-Host "                  maximum entropy, harder to communicate verbally" -ForegroundColor DarkGray
    Write-Host ""
    $StyleChoice = (Read-Host "Enter choice (1 or 2 - press Enter for 1)").Trim()
    if ([string]::IsNullOrWhiteSpace($StyleChoice)) { $StyleChoice = '1' }
    $GenStyle = if ($StyleChoice -eq '2') { 'Random' } else { 'Readable' }

    if ($GenStyle -eq 'Random') {
        Write-Host ""
        $LenInput = (Read-Host "Password length 12-64 (press Enter for 14)").Trim()
        if (-not [string]::IsNullOrWhiteSpace($LenInput)) {
            $parsed = 0
            if ([int]::TryParse($LenInput, [ref]$parsed)) { $GenLength = $parsed }
            else { Write-Host "[WARN] Not a number - using 14." -ForegroundColor DarkYellow }
        }
        if ($GenLength -lt 12) { $GenLength = 12; Write-Host "[WARN] Clamped up to 12." -ForegroundColor DarkYellow }
        if ($GenLength -gt 64) { $GenLength = 64; Write-Host "[WARN] Clamped down to 64." -ForegroundColor DarkYellow }
    }
    else {
        Write-Host ""
        Write-Host "Word count (more words = more entropy, longer to dictate):" -ForegroundColor Cyan
        Write-Host "  [2] Word-Noun-####!        approx. $(Get-PasswordEntropyBits -Style Readable -WordCount 2 -DigitCount 4) bits" -ForegroundColor White
        Write-Host "  [3] Word-Word-Noun-####!   approx. $(Get-PasswordEntropyBits -Style Readable -WordCount 3 -DigitCount 4) bits  (default)" -ForegroundColor White
        Write-Host "  [4] Word-Word-Word-Noun-####!  approx. $(Get-PasswordEntropyBits -Style Readable -WordCount 4 -DigitCount 4) bits" -ForegroundColor White
        Write-Host ""
        $WcInput = (Read-Host "Enter 2, 3, or 4 (press Enter for 3)").Trim()
        if ($WcInput -in '2','3','4') { $GenWordCount = [int]$WcInput }
    }

    $GenEntropy = Get-PasswordEntropyBits -Style $GenStyle -Length $GenLength -WordCount $GenWordCount -DigitCount $GenDigitCount

    Write-Host ""
    Write-Host "Generator configured:" -ForegroundColor Green
    Write-Host "  Style   : $GenStyle" -ForegroundColor Green
    if ($GenStyle -eq 'Random') {
        Write-Host "  Length  : $GenLength chars from a 64-char unambiguous pool" -ForegroundColor Green
    } else {
        Write-Host "  Pattern : $GenWordCount words + $GenDigitCount digits + 1 symbol, '-' separated" -ForegroundColor Green
        Write-Host "  Wordlist: $($PwdWordsA.Count) modifiers x $($PwdWordsB.Count) nouns (embedded, offline)" -ForegroundColor DarkGray
    }
    Write-Host "  Entropy : ~$GenEntropy bits" -ForegroundColor Green
    Write-Host "  RNG     : System.Security.Cryptography.RandomNumberGenerator (CSPRNG)" -ForegroundColor DarkGray
    Write-Host "  Samples : $(New-ReadablePassword -WordCount $GenWordCount -DigitCount $GenDigitCount)" -ForegroundColor DarkGray
    if ($GenEntropy -lt 45) {
        Write-Host "  [NOTE] Under 45 bits. Acceptable for a forced-change-at-next-sign-in temporary" -ForegroundColor DarkYellow
        Write-Host "         credential protected by Entra smart lockout. Not for permanent passwords." -ForegroundColor DarkYellow
    }
}

# ---------- 4d. Run options: progress heartbeat + force sign-out ----------
Write-Host ""
Write-Host "--- Run options ---" -ForegroundColor Cyan

# Heartbeat: useful on 500-row batches, pure noise on 20-row batches.
Write-Host ""
Write-Host "Progress heartbeat prints a running OK/FAIL/SKIP tally every N rows." -ForegroundColor Gray
$HbInput = (Read-Host "Heartbeat every N rows - enter 0 to switch it OFF (press Enter for 25)").Trim()
$HeartbeatEvery = 25
if (-not [string]::IsNullOrWhiteSpace($HbInput)) {
    $hbParsed = 0
    if ([int]::TryParse($HbInput, [ref]$hbParsed) -and $hbParsed -ge 0) { $HeartbeatEvery = $hbParsed }
    else { Write-Host "[WARN] Not a valid number - keeping 25." -ForegroundColor DarkYellow }
}
if ($HeartbeatEvery -eq 0) {
    Write-Host "  Heartbeat: OFF (per-row lines only)" -ForegroundColor DarkGray
} else {
    Write-Host "  Heartbeat: every $HeartbeatEvery rows" -ForegroundColor DarkGray
}

# Force sign-out. Runs AFTER the password is changed, so the new password is the only way back in.
$RevokeSessions = $false
Write-Host ""
if (-not $RevokeCapable) {
    Write-Host "Force sign-out: UNAVAILABLE (required scope not granted - see Phase 2)." -ForegroundColor DarkYellow
}
else {
    Write-Host "Force sign-out (Revoke-MgUserSignInSession) after each successful reset?" -ForegroundColor Cyan
    Write-Host "  Invalidates all refresh tokens and browser session cookies by stamping" -ForegroundColor DarkGray
    Write-Host "  signInSessionsValidFromDateTime = now. Signs the user out of Outlook, Teams," -ForegroundColor DarkGray
    Write-Host "  OneDrive and the web everywhere." -ForegroundColor DarkGray
    Write-Host "  CAVEAT: already-issued ACCESS tokens stay valid until they expire (up to ~1 hour)," -ForegroundColor DarkYellow
    Write-Host "  so this is not instant. Services using Continuous Access Evaluation react sooner." -ForegroundColor DarkYellow
    Write-Host "  Microsoft also notes a few minutes of propagation delay." -ForegroundColor DarkYellow
    Write-Host ""
    $revokeAns = (Read-Host "Force sign-out after reset? (Y/N - press Enter for N)").Trim()
    $RevokeSessions = ($revokeAns -match '^(y|yes)$')
    if ($RevokeSessions) {
        Write-Host "  Force sign-out: ENABLED - runs only on rows whose reset SUCCEEDS." -ForegroundColor Yellow
    } else {
        Write-Host "  Force sign-out: disabled" -ForegroundColor DarkGray
    }
}

# Privileged-account scan. This is the slowest read in the script on a large tenant, so it
# gets an explicit opt-out rather than silently making the operator wait.
Write-Host ""
Write-Host "Scan the tenant for privileged accounts before resetting?" -ForegroundColor Cyan
Write-Host "  Lists members of Global Administrator, Privileged Authentication Administrator," -ForegroundColor DarkGray
Write-Host "  User Administrator and Privileged Role Administrator, so the script can warn you" -ForegroundColor DarkGray
Write-Host "  and demand a typed CONFIRM before resetting any admin account." -ForegroundColor DarkGray
Write-Host "  Usually a few seconds; can take longer on large tenants." -ForegroundColor DarkGray
$scanAns = (Read-Host "Scan for privileged accounts? (Y/N - press Enter for Y)").Trim()
$ScanPrivileged = -not ($scanAns -match '^(n|no)$')

# Write mode. Sequential is one Graph call per user (~2-3/sec). Batched sends 20 PATCHes per
# call, executed in parallel server-side, and paces itself just under the Entra write ceiling.
$UseBatchWrites = $false
Write-Host ""
Write-Host "Write mode:" -ForegroundColor Cyan
Write-Host "  [1] Sequential - one Graph call per user, ~2-3 writes/sec" -ForegroundColor White
Write-Host "  [2] Batched    - $GraphBatchMax PATCHes per Graph call, paced to ~$TargetWritesPerSec writes/sec" -ForegroundColor White
Write-Host ""
if ($ProcessedRows.Count -ge 100) {
    $seqEst = Format-Duration ($ProcessedRows.Count * 0.45)
    $batEst = Format-Duration ($ProcessedRows.Count / $TargetWritesPerSec)
    Write-Host "  For $($ProcessedRows.Count) row(s): sequential ~$seqEst vs batched ~$batEst." -ForegroundColor Yellow
    Write-Host "  Batched is recommended at this size." -ForegroundColor Yellow
    Write-Host "  NOTE: in batched mode the audit log is written in EXECUTION order." -ForegroundColor DarkGray
    Write-Host "        Sort by the Row column for original CSV order." -ForegroundColor DarkGray
    $defaultWrite = '2'
} else {
    Write-Host "  For $($ProcessedRows.Count) row(s) batching saves little; sequential is simpler to read." -ForegroundColor DarkGray
    $defaultWrite = '1'
}
$writeAns = (Read-Host "Enter choice (1 or 2 - press Enter for $defaultWrite)").Trim()
if ([string]::IsNullOrWhiteSpace($writeAns)) { $writeAns = $defaultWrite }
$UseBatchWrites = ($writeAns -eq '2')
if ($UseBatchWrites) {
    Write-Host "  Write mode: BATCHED ($GraphBatchMax per call, governor targeting $TargetWritesPerSec writes/sec)" -ForegroundColor Green
} else {
    Write-Host "  Write mode: sequential" -ForegroundColor DarkGray
}

# Audit-log flush cadence. Append-only since v3.5.0, so this stays cheap at any batch size.
$FlushEvery = if ($UseBatchWrites) { $GraphBatchMax } else { 10 }
if ($ScanPrivileged) {
    Write-Host "  Privileged scan: ON (admin-account gate active)" -ForegroundColor DarkGray
} else {
    Write-Host "  Privileged scan: OFF - admin accounts will NOT be flagged or gated." -ForegroundColor DarkYellow
}

# ---------- 4e. Pre-flight against Entra ----------
Write-Host ""
Write-Host "Running pre-flight checks against Entra ID..." -ForegroundColor Cyan

$PrivilegedUpns = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

if (-not $ScanPrivileged) {
    Write-Host "  Privileged-account scan SKIPPED by choice - the admin-account gate is OFF." -ForegroundColor DarkYellow
}
else {
try {
    $PrivilegedRoleNames = @('Global Administrator','Privileged Authentication Administrator','User Administrator','Privileged Role Administrator')
    $privScanStart = Get-Date

    Write-Host "  Listing activated directory roles..." -ForegroundColor DarkGray
    $ActiveRoles  = @(Invoke-GraphWithRetry -ScriptBlock { Get-MgDirectoryRole -All -ErrorAction Stop })
    $RoleTargets  = @($ActiveRoles | Where-Object { $_.DisplayName -in $PrivilegedRoleNames })
    Write-Host "  $($ActiveRoles.Count) activated role(s); $($RoleTargets.Count) are privileged and will be scanned." -ForegroundColor DarkGray

    # Per-role progress. Each of these is a paged query and can take seconds on a large
    # tenant; without output the whole block looks like a freeze.
    $ri = 0
    foreach ($role in $RoleTargets) {
        $ri++
        Write-Host ("    [{0}/{1}] {2} ..." -f $ri, $RoleTargets.Count, $role.DisplayName) -NoNewline -ForegroundColor DarkGray
        try {
            $roleStart = Get-Date
            $members = @(Invoke-GraphWithRetry -ScriptBlock {
                Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -All -ErrorAction Stop
            })
            $added = 0
            foreach ($m in $members) {
                try {
                    $upnProp = $m.AdditionalProperties['userPrincipalName']
                    if ($upnProp -and $PrivilegedUpns.Add([string]$upnProp)) { $added++ }
                } catch { }
            }
            Write-Host (" {0} member(s), {1} new ({2:N1}s)" -f $members.Count, $added, ((Get-Date) - $roleStart).TotalSeconds) -ForegroundColor DarkGray
        }
        catch {
            Write-Host " FAILED: $($_.Exception.Message)" -ForegroundColor DarkYellow
        }
    }

    Write-Host ("  Loaded {0} privileged account(s) for protection check ({1:N1}s total)." -f `
        $PrivilegedUpns.Count, ((Get-Date) - $privScanStart).TotalSeconds) -ForegroundColor DarkGray
} catch {
    $roleErr = [string]$_.Exception.Message
    if (Test-IsAuthFailure -Message $roleErr) {
        # Token died between Phase 2 and here - degrading to "no admin protection" would be
        # unsafe, and every row is about to fail anyway.
        Write-Host ""
        Write-Host "[FAIL] The Graph token stopped working: $roleErr" -ForegroundColor Red
        Write-Host "       This is an authentication failure, not a data problem, so admin-account" -ForegroundColor Red
        Write-Host "       protection cannot be trusted and every reset would fail." -ForegroundColor Red
        Write-Host "       Reconnect (sign-in mode [1] Interactive) and re-run." -ForegroundColor Yellow
        Write-Host "Aborting - nothing was changed, no files written." -ForegroundColor Red
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        return
    }
    Write-Host "  [WARN] Could not enumerate privileged roles: $roleErr" -ForegroundColor DarkYellow
    Write-Host "  [WARN] Continuing WITHOUT admin-account protection." -ForegroundColor DarkYellow
}
}

Write-Host ""
Write-Host "Resolving $($ProcessedRows.Count) user(s) against the tenant..." -ForegroundColor Cyan
$LookupStart = Get-Date

# Choose the lookup strategy by batch size. One request per row is cheapest for small CSVs;
# past a few hundred rows a paged bulk read of the directory wins by orders of magnitude.
$BulkThreshold = 200
$UseBulkIndex  = ($ProcessedRows.Count -ge $BulkThreshold)
$TenantIndex   = $null

if ($UseBulkIndex) {
    $perRowReqs = $ProcessedRows.Count
    $bulkReqs   = [Math]::Ceiling($ProcessedRows.Count / 999)
    Write-Host "  Large batch: using a bulk directory read instead of one lookup per user." -ForegroundColor Cyan
    Write-Host "  (one-by-one would be ~$perRowReqs requests; paged bulk read is ~$bulkReqs)" -ForegroundColor DarkGray
    Write-Host "  Downloading and indexing the directory..." -ForegroundColor DarkGray
    try {
        $TenantIndex = Get-TenantUserIndex
        Write-Host ("  Indexed {0} tenant user(s) in {1}." -f $TenantIndex.Count, (Format-Duration ((Get-Date) - $LookupStart).TotalSeconds)) -ForegroundColor Green
    }
    catch {
        $bulkErr = [string]$_.Exception.Message
        if (Test-IsAuthFailure -Message $bulkErr) {
            Write-Host "[FAIL] Authentication failed during the bulk directory read: $bulkErr" -ForegroundColor Red
            Write-Host "       Reconnect (sign-in mode [1] Interactive) and re-run." -ForegroundColor Yellow
            Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
            return
        }
        Write-Host "  [WARN] Bulk directory read failed: $bulkErr" -ForegroundColor DarkYellow
        Write-Host "  [WARN] Falling back to one lookup per user. This will be MUCH slower." -ForegroundColor DarkYellow
        $UseBulkIndex = $false
        $TenantIndex  = $null
    }
}
else {
    Write-Host "  Small batch: one lookup per user." -ForegroundColor DarkGray
}
$MatchStart = Get-Date

$PreflightCounts = @{
    Found = 0; NotFound = 0; Hybrid = 0; Disabled = 0; Guest = 0
    Privileged = 0; InvalidUpn = 0; DuplicateSkip = 0; BadPassword = 0; ResumeSkip = 0
}
$SeenUpns = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$Checked  = 0
$PreflightAuthFailure = ''

foreach ($row in $ProcessedRows) {
    $Checked++

    if (-not $row.UpnValid) {
        $row.PreflightStatus  = 'Skip'
        $row.PreflightMessage = "Invalid UPN: $($row.UpnReason)"
        $PreflightCounts.InvalidUpn++
        continue
    }
    if (-not $SeenUpns.Add($row.UpnClean)) {
        $row.PreflightStatus  = 'Skip'
        $row.PreflightMessage = 'Duplicate UPN in CSV (earlier row wins)'
        $PreflightCounts.DuplicateSkip++
        continue
    }
    # Resume: already done in a previous run. Checked BEFORE the tenant lookup so a resumed
    # run costs nothing for work it is skipping, and no password is generated for it.
    if ($ResumeSkipUpns.Count -gt 0 -and $ResumeSkipUpns.Contains($row.UpnClean)) {
        $row.PreflightStatus  = 'Skip'
        $row.PreflightMessage = 'Already reset successfully in the resume log'
        $PreflightCounts.ResumeSkip++
        continue
    }

    try {
        if ($UseBulkIndex) {
            # In-memory hit - no network, no throttling, no per-row latency.
            if ($TenantIndex.ContainsKey($row.UpnClean)) {
                $mgUser = $TenantIndex[$row.UpnClean]
            } else {
                throw 'Not present in the tenant directory index'
            }
        }
        else {
            $mgUser = Invoke-GraphWithRetry -ScriptBlock {
                Get-MgUser -UserId $row.UpnClean -Property 'Id,UserPrincipalName,AccountEnabled,UserType,OnPremisesSyncEnabled,DisplayName' -ErrorAction Stop
            }
        }
        $row.UserObjectId = [string]$mgUser.Id
        $row.DisplayName  = [string]$mgUser.DisplayName

        $notes = @()
        if ($mgUser.OnPremisesSyncEnabled -eq $true) { $PreflightCounts.Hybrid++;   $notes += 'HYBRID (on-prem synced - reset would fail)' }
        if ($mgUser.AccountEnabled -eq $false)       { $PreflightCounts.Disabled++; $notes += 'Disabled' }
        if ($mgUser.UserType -eq 'Guest')            { $PreflightCounts.Guest++;    $notes += 'Guest' }
        if ($PrivilegedUpns.Contains([string]$mgUser.UserPrincipalName)) {
            $PreflightCounts.Privileged++; $notes += 'PRIVILEGED ROLE HOLDER'
        }

        if ($mgUser.OnPremisesSyncEnabled -eq $true) {
            $row.PreflightStatus  = 'Skip'
            $row.PreflightMessage = ($notes -join '; ')
        } else {
            $row.PreflightStatus  = 'Ready'
            $row.PreflightMessage = if ($notes.Count -gt 0) { ($notes -join '; ') } else { 'OK' }
        }
    }
    catch {
        $lookupErr = [string]$_.Exception.Message

        # Do NOT record an auth failure as "user not found" - that would silently convert a
        # broken session into 26 bogus "missing user" rows and a misleading dry-run report.
        if (Test-IsAuthFailure -Message $lookupErr) {
            $PreflightAuthFailure = $lookupErr
            break
        }

        $row.PreflightStatus  = 'Skip'
        $row.PreflightMessage = "Not found in tenant: $lookupErr"
        $PreflightCounts.NotFound++
    }

    if ($HeartbeatEvery -gt 0 -and ($Checked % $HeartbeatEvery) -eq 0) {
        Write-Host ("  [HEARTBEAT] Pre-flight {0}/{1}..." -f $Checked, $ProcessedRows.Count) -ForegroundColor DarkGray
    }
}

if (-not $PreflightAuthFailure) {
    $matchSecs = ((Get-Date) - $MatchStart).TotalSeconds
    $totalSecs = ((Get-Date) - $LookupStart).TotalSeconds
    if ($UseBulkIndex) {
        Write-Host ("  Matched {0} row(s) against the index in {1} (total pre-flight {2})." -f `
            $Checked, (Format-Duration $matchSecs), (Format-Duration $totalSecs)) -ForegroundColor DarkGray
    } else {
        Write-Host ("  Resolved {0} user(s) in {1} ({2:N2}s/user)." -f `
            $Checked, (Format-Duration $totalSecs), ($totalSecs / [Math]::Max(1, $Checked))) -ForegroundColor DarkGray
    }
}

if ($PreflightAuthFailure) {
    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Red
    Write-Host " ABORTED - AUTHENTICATION FAILURE DURING PRE-FLIGHT" -ForegroundColor Red
    Write-Host "============================================================" -ForegroundColor Red
    Write-Host " $PreflightAuthFailure" -ForegroundColor DarkRed
    Write-Host ""
    Write-Host " Stopped at row $Checked of $($ProcessedRows.Count). This is your SESSION, not your data -" -ForegroundColor Yellow
    Write-Host " the remaining users were never checked and none of them are missing." -ForegroundColor Yellow
    if ($PreflightAuthFailure -match 'DeviceCodeCredential|Object reference not set to an instance') {
        Write-Host ""
        Write-Host " Known Graph SDK 2.34+ device-code regression (issue #3495)." -ForegroundColor Cyan
        Write-Host " Re-run and choose sign-in mode [1] Interactive popup." -ForegroundColor Cyan
    } else {
        Write-Host ""
        Write-Host " Reconnect and re-run. If it recurs, check for Microsoft.Graph.* version drift." -ForegroundColor Cyan
    }
    Write-Host " Nothing was changed and no credentials file was written." -ForegroundColor Green
    Write-Host "============================================================" -ForegroundColor Red
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    return
}

# ---------- 4f. Resolve the password for every Ready row ----------
Write-Host ""
Write-Host "Resolving passwords (mode: $PasswordMode)..." -ForegroundColor Cyan

$IssuedPasswords = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
$SourceCounts    = @{ CSV = 0; Generated = 0 }
$GenerationError = $null

foreach ($row in $ProcessedRows) {
    if ($row.PreflightStatus -ne 'Ready') { continue }

    $useCsv = switch ($PasswordMode) {
        'CsvOnly'     { $true }
        'GenerateAll' { $false }
        'Hybrid'      { [bool]$row.CsvPassword }
    }

    if ($useCsv) {
        $row.Password       = $row.CsvPassword
        $row.PasswordSource = 'CSV'
        [void]$IssuedPasswords.Add($row.CsvPassword)
    }
    else {
        try {
            $row.Password       = New-UniquePassword -UsedSet $IssuedPasswords -Style $GenStyle `
                                                     -Length $GenLength -WordCount $GenWordCount -DigitCount $GenDigitCount
            $row.PasswordSource = 'Generated'
        }
        catch {
            $GenerationError = $_.Exception.Message
            break
        }
    }

    # One policy definition, applied to CSV and generated passwords alike
    $row.PasswordIssues = Test-PasswordPolicy -Password $row.Password -Upn $row.UpnClean
    if ($row.PasswordIssues) {
        $row.PreflightStatus  = 'Skip'
        $row.PreflightMessage = "Password issue ($($row.PasswordSource)): $($row.PasswordIssues)"
        $PreflightCounts.BadPassword++
        continue
    }

    if ($row.PasswordSource -eq 'CSV') { $SourceCounts.CSV++ } else { $SourceCounts.Generated++ }
    $PreflightCounts.Found++
}

if ($GenerationError) {
    Write-Host "[FAIL] Password generation failed: $GenerationError" -ForegroundColor Red
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    return
}

# ---------- 4g. Summary ----------
Write-Host ""
Write-Host "Pre-flight summary:" -ForegroundColor Cyan
Write-Host ("  Ready to process     : {0}" -f $PreflightCounts.Found)         -ForegroundColor Green
Write-Host ("     from CSV password : {0}" -f $SourceCounts.CSV)              -ForegroundColor DarkGreen
Write-Host ("     generated         : {0}" -f $SourceCounts.Generated)        -ForegroundColor DarkGreen
Write-Host ("  Not found in tenant  : {0}" -f $PreflightCounts.NotFound)      -ForegroundColor Red
Write-Host ("  Hybrid (blocked)     : {0}" -f $PreflightCounts.Hybrid)        -ForegroundColor Red
Write-Host ("  Disabled accounts    : {0}" -f $PreflightCounts.Disabled)      -ForegroundColor Yellow
Write-Host ("  Guest accounts       : {0}" -f $PreflightCounts.Guest)         -ForegroundColor Yellow
Write-Host ("  Privileged (admin)   : {0}" -f $PreflightCounts.Privileged)    -ForegroundColor Yellow
Write-Host ("  Invalid UPN format   : {0}" -f $PreflightCounts.InvalidUpn)    -ForegroundColor Red
Write-Host ("  Password issues      : {0}" -f $PreflightCounts.BadPassword)   -ForegroundColor Red
Write-Host ("  Duplicates in CSV    : {0}" -f $PreflightCounts.DuplicateSkip) -ForegroundColor Yellow
Write-Host ("  Skipped by resume    : {0}" -f $PreflightCounts.ResumeSkip)    -ForegroundColor Yellow
Write-Host ("  UPNs sanitized       : {0}" -f $SanitizedCount)                -ForegroundColor DarkGray

if ($IsDryRun) {
    Write-Host ""
    Write-Host "DRY-RUN MODE: no passwords will be changed and NO credentials file will be written." -ForegroundColor Cyan
    if ($RevokeSessions) {
        Write-Host "  Force sign-out was requested: $($PreflightCounts.Found) user(s) WOULD be signed out on a real run." -ForegroundColor Yellow
        Write-Host "  No sessions were revoked in dry-run." -ForegroundColor DarkGray
    }
    $DryPath = Join-Path $LogFolder "ResetPasswords_DryRun_$TimeStamp.csv"
    $ProcessedRows | Select-Object Row, UpnRaw, UpnClean, UpnChanged, UpnValid, UserObjectId, DisplayName,
        PasswordSource,
        @{n='PasswordLength';e={ $_.Password.Length }},
        @{n='PasswordMasked';e={ Format-MaskedSecret $_.Password }},
        PasswordIssues, PreflightStatus, PreflightMessage,
        @{n='CorrelationId';e={ $CorrelationId }} |
        Export-Csv -LiteralPath $DryPath -NoTypeInformation -Encoding UTF8
    Write-Host "  [OK] Dry-run report (masked, no secrets): $DryPath" -ForegroundColor Green
    Write-Host "  Re-run in mode [Y] to generate real credentials and apply the resets." -ForegroundColor DarkGray
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    return
}

# ---------- 4h. Confirmation gates ----------
$ReadyCount = $PreflightCounts.Found
if ($ReadyCount -eq 0) {
    Write-Host ""
    Write-Host "Nothing to process. Exiting." -ForegroundColor Red
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    return
}

Write-Host ""
Write-Host "About to RESET PASSWORDS for $ReadyCount user(s) in tenant $TenantId." -ForegroundColor Yellow
Write-Host "  $($SourceCounts.CSV) from the CSV, $($SourceCounts.Generated) newly generated." -ForegroundColor Yellow
Write-Host "Each user will be required to change their password at next sign-in." -ForegroundColor Yellow
if ($RevokeSessions) {
    Write-Host ""
    Write-Host "!! FORCE SIGN-OUT IS ENABLED !!" -ForegroundColor Red
    Write-Host "   Every successfully reset user will also be signed out of Outlook, Teams," -ForegroundColor Yellow
    Write-Host "   OneDrive and all browser sessions on every device. Expect helpdesk calls." -ForegroundColor Yellow
    if ($OperatorAccount -in @($ProcessedRows | Where-Object { $_.PreflightStatus -eq 'Ready' }).UpnClean) {
        Write-Host "   [PROTECTED] Your own account ($OperatorAccount) is in this batch. Its password" -ForegroundColor Cyan
        Write-Host "               will be reset but sign-out will be SKIPPED, so this run can finish." -ForegroundColor Cyan
    }
}

if (-not $ScanPrivileged) {
    Write-Host ""
    Write-Host "[NOTE] Privileged scan was skipped, so admin accounts in this batch were NOT" -ForegroundColor DarkYellow
    Write-Host "       detected and the CONFIRM gate will not fire. Proceed only if you are sure" -ForegroundColor DarkYellow
    Write-Host "       this CSV contains no administrator accounts." -ForegroundColor DarkYellow
}

if ($PreflightCounts.Privileged -gt 0) {
    Write-Host ""
    Write-Host "!! WARNING: $($PreflightCounts.Privileged) target account(s) hold privileged roles !!" -ForegroundColor Red
    $adminConfirm = Read-Host "Type CONFIRM to include admin accounts in the reset"
    if ($adminConfirm -ne 'CONFIRM') {
        Write-Host "Not confirmed. Cancelling." -ForegroundColor Red
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        return
    }
}

if ($ReadyCount -gt 500) {
    Write-Host ""
    Write-Host "Large batch detected ($ReadyCount rows)." -ForegroundColor Yellow
    $sizeConfirm = Read-Host "Type the exact count '$ReadyCount' to proceed"
    if ($sizeConfirm -ne $ReadyCount.ToString()) {
        Write-Host "Count mismatch. Cancelling." -ForegroundColor Red
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        return
    }
}

# Duration expectation. Password writes are one Graph call each and cannot be batched, so a
# very large batch is a multi-hour job. Say so BEFORE the operator commits, not at hour three.
if ($ReadyCount -ge 250) {
    if ($UseBatchWrites) {
        $estLow  = Format-Duration ($ReadyCount / 20.0)
        $estHigh = Format-Duration ($ReadyCount / 12.0)
    } else {
        $estLow  = Format-Duration ($ReadyCount * 0.30)
        $estHigh = Format-Duration ($ReadyCount * 0.70)
    }
    Write-Host ""
    Write-Host "Expected execution time: $estLow to $estHigh" -ForegroundColor Cyan
    if ($UseBatchWrites) {
        Write-Host "  $ReadyCount resets in $([Math]::Ceiling($ReadyCount / $GraphBatchMax)) batched Graph call(s) of up to $GraphBatchMax," -ForegroundColor DarkGray
        Write-Host "  paced to ~$TargetWritesPerSec writes/sec (Entra allows ~20/sec per app per tenant)." -ForegroundColor DarkGray
    } else {
        Write-Host "  $ReadyCount resets, one Graph write each, at roughly 1.5-3 writes/sec." -ForegroundColor DarkGray
        Write-Host "  Batched write mode would cut this to about $(Format-Duration ($ReadyCount / 15.0))." -ForegroundColor DarkGray
    }
    Write-Host "  Keep this window open. A live ETA appears in the heartbeat." -ForegroundColor DarkGray
    if ($HeartbeatEvery -eq 0) {
        Write-Host "  [WARN] Heartbeat is OFF for a run this long - you will get no ETA and no" -ForegroundColor DarkYellow
        Write-Host "         running tally. Consider re-running with a heartbeat of 100-250." -ForegroundColor DarkYellow
    }
    Write-Host "  Safe to interrupt: every password is already recorded in the credentials file," -ForegroundColor DarkGray
    Write-Host "  and the audit log is flushed every $FlushEvery rows." -ForegroundColor DarkGray
}

Write-Host ""
$Confirm = Read-Host "Final confirmation - proceed with $ReadyCount password reset(s)? (Y/N)"
if ($Confirm -notmatch '^(y|yes)$') {
    Write-Host "Cancelled. No changes were made." -ForegroundColor Red
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    return
}

# ---------- 4i. PRE-WRITE the credentials file (before any tenant change) ----------
# This is the crash-safety guarantee. If PowerShell dies at row 87, every password
# that was or will be applied is already on disk, so nobody gets locked out.
Write-Host ""
Write-Host "Writing credentials file BEFORE any reset is applied..." -ForegroundColor Cyan

$CredRows = @($ProcessedRows | Where-Object { $_.PreflightStatus -eq 'Ready' } |
    Select-Object @{n='upn';e={ $_.UpnClean }},
                  @{n='displayName';e={ $_.DisplayName }},
                  @{n='password';e={ $_.Password }},
                  @{n='source';e={ $_.PasswordSource }},
                  @{n='mustChangeAtNextSignIn';e={ 'Yes' }},
                  @{n='correlationId';e={ $CorrelationId }})

try {
    $CredRows | Export-Csv -LiteralPath $CredentialsPath -NoTypeInformation -Encoding UTF8 -Force -ErrorAction Stop
}
catch {
    Write-Host "[FAIL] Could not write the credentials file: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "ABORTING before any password is changed - a reset without a recorded password" -ForegroundColor Red
    Write-Host "would lock users out permanently." -ForegroundColor Red
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    return
}

# Verify by reading it back - a silent partial write is the exact failure we are guarding against
try {
    $Verify = @(Import-Csv -LiteralPath $CredentialsPath -Encoding UTF8 -ErrorAction Stop)
    if ($Verify.Count -ne $CredRows.Count) {
        throw "Read-back row count $($Verify.Count) does not match expected $($CredRows.Count)."
    }
}
catch {
    Write-Host "[FAIL] Credentials file failed verification: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "ABORTING before any password is changed." -ForegroundColor Red
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    return
}

$AclLocked = Set-SecretFileAcl -Path $CredentialsPath
Write-Host "  [OK] Credentials file written and verified ($($CredRows.Count) rows)." -ForegroundColor Green
Write-Host "       $CredentialsPath" -ForegroundColor Yellow
if ($AclLocked) {
    Write-Host "  [OK] NTFS permissions restricted to $env:USERDOMAIN\$env:USERNAME." -ForegroundColor Green
} else {
    Write-Host "  [WARN] Could not restrict NTFS permissions - protect this file manually." -ForegroundColor DarkYellow
}
Write-Host "  [ACTION] Distribute the passwords, then DELETE this file." -ForegroundColor Yellow

Write-Host ""
Write-Host "Confirmed. Proceeding..." -ForegroundColor Green


# ============================================================
# PHASE 5 - Execute + Flush + Export Audit
# ============================================================
Write-PhaseHeader "PHASE 5: Reset Passwords + Export Audit Log"

$RunStart   = Get-Date
$Log        = [System.Collections.Generic.List[object]]::new()
$Counters   = @{ Success = 0; Failed = 0; Skipped = 0; Retried = 0; Revoked = 0; RevokeFailed = 0; RevokeSkipped = 0 }

# Force sign-out is a bonus action, never allowed to fail the reset it follows.
# If it 403s repeatedly the scope is wrong, so switch it off and let the resets continue.
$RevokeAuthFails       = 0
$RevokeDisabledMidRun  = $false
$RowIndex   = 0
$Total      = $ProcessedRows.Count
$FlushedRows = 0     # append-only cursor into $Log
if (-not (Get-Variable -Name FlushEvery -Scope 0 -ErrorAction SilentlyContinue)) { $FlushEvery = 10 }

# Circuit breaker. A tenant-wide authorization problem fails EVERY row identically; there is
# no point issuing 500 more doomed calls, and each one burns a Graph resource unit.
$AuthFailThreshold    = 5
$ConsecutiveAuthFails = 0
$AbortReason          = ''

$AuditCommon = @{
    CorrelationId = $CorrelationId
    Operator      = $OperatorAccount
    Tenant        = $TenantId
    Version       = $ScriptVersion
    GenStyle      = $GenStyle
}

# ============================================================
# PHASE 5A - BATCHED WRITE PATH
# ============================================================
if ($UseBatchWrites) {

    # Skips are instant and need no Graph call - log them first, in row order.
    foreach ($row in $ProcessedRows | Where-Object { $_.PreflightStatus -ne 'Ready' }) {
        $Counters.Skipped++
        $Log.Add((New-AuditRow @AuditCommon -Row $row -Status 'Skipped' `
                    -Message $row.PreflightMessage -WriteMode 'Batched'))
    }
    Write-Host "Logged $($Counters.Skipped) skipped row(s)." -ForegroundColor Yellow

    $WorkRows  = @($ProcessedRows | Where-Object { $_.PreflightStatus -eq 'Ready' })
    $ChunkCount = [Math]::Ceiling($WorkRows.Count / [double]$GraphBatchMax)
    Write-Host "Writing $($WorkRows.Count) reset(s) in $ChunkCount batch call(s) of up to $GraphBatchMax..." -ForegroundColor Cyan

    # Governor state. Two jobs: pace proactively so we sit under ~20 writes/sec, and back off
    # hard when Graph tells us we are over.
    $GovDelayMs    = 0
    $GovCleanRuns  = 0
    $GovThrottled  = 0
    $ChunkIndex    = 0

    for ($offset = 0; $offset -lt $WorkRows.Count; $offset += $GraphBatchMax) {
        $ChunkIndex++
        $chunk      = @($WorkRows[$offset..([Math]::Min($offset + $GraphBatchMax - 1, $WorkRows.Count - 1))])
        $chunkStart = Get-Date

        # --- build the reset sub-requests; index by sub-request id ---
        $byId    = @{}
        $requests = @()
        for ($i = 0; $i -lt $chunk.Count; $i++) {
            $sid = "r$i"
            $byId[$sid] = $chunk[$i]
            # Target by object id, not UPN - no URL-escaping surprises with odd UPNs.
            $requests += @{
                id      = $sid
                method  = 'PATCH'
                url     = "/users/$($chunk[$i].UserObjectId)"
                headers = @{ 'Content-Type' = 'application/json' }
                body    = @{ passwordProfile = @{
                                password                      = $chunk[$i].Password
                                forceChangePasswordNextSignIn = $true } }
            }
        }

        $resetResults = $null
        $batchFatal   = ''
        try {
            $batch        = Invoke-GraphBatch -Requests $requests
            $resetResults = $batch.Results
            if ($batch.Throttled) { $GovThrottled++ }
        }
        catch {
            $batchFatal = [string]$_.Exception.Message
            if (Test-IsAuthFailure -Message $batchFatal) {
                $AbortReason = "Authentication failed during batched write: $batchFatal"
                $FlushedRows = Save-LogIncremental -Log $Log -Path $LogPath -AlreadyFlushed $FlushedRows
                Write-Host ""
                Write-Host "[FAIL] $AbortReason" -ForegroundColor Red
                Write-Host "       Stopped at batch $ChunkIndex of $ChunkCount. Reconnect and resume using" -ForegroundColor Yellow
                Write-Host "       this run's audit log as the resume source." -ForegroundColor Yellow
                break
            }
        }

        # --- classify each row in the chunk ---
        $chunkSuccess = @()
        $rowOutcomes  = @{}
        foreach ($sid in $byId.Keys) {
            $row = $byId[$sid]
            if ($batchFatal) {
                $rowOutcomes[$sid] = @{ Status='Failed'; Message="Batch call failed: $batchFatal"; Http=''; Code='BatchTransportFailure' }
                continue
            }
            $res = if ($resetResults.ContainsKey($sid)) { $resetResults[$sid] } else { $null }
            if ($null -eq $res) {
                $rowOutcomes[$sid] = @{ Status='Failed'; Message='No response returned for this sub-request'; Http=''; Code='MissingBatchResponse' }
                continue
            }
            if ($res.Status -ge 200 -and $res.Status -lt 300) {
                $rowOutcomes[$sid] = @{ Status='Success'; Message='Password reset successful'; Http=[string]$res.Status; Code='' }
                $chunkSuccess += $sid
            }
            else {
                $errText = Get-BatchErrorText -Body $res.Body
                $code    = ''
                if ($errText -match '^\[([^\]]+)\]') { $code = $Matches[1] }
                $rowOutcomes[$sid] = @{ Status='Failed'; Message=$errText; Http=[string]$res.Status; Code=$code }
            }
        }

        # --- force sign-out for this chunk's successes, batched too ---
        $revokeOutcomes = @{}
        if ($RevokeSessions -and -not $RevokeDisabledMidRun -and $chunkSuccess.Count -gt 0) {
            $revReqs = @()
            foreach ($sid in $chunkSuccess) {
                $row = $byId[$sid]
                if ($row.UpnClean -eq $OperatorAccount) {
                    $revokeOutcomes[$sid] = @{ Status='Skipped'; Message='Operator account - sign-out skipped to keep this run alive' }
                    $Counters.RevokeSkipped++
                    continue
                }
                $revReqs += @{
                    id      = "v$sid"
                    method  = 'POST'
                    url     = "/users/$($row.UserObjectId)/revokeSignInSessions"
                    headers = @{ 'Content-Type' = 'application/json' }
                    body    = @{}
                }
            }
            if ($revReqs.Count -gt 0) {
                try {
                    $rBatch = Invoke-GraphBatch -Requests $revReqs
                    $rAuth  = 0
                    foreach ($sid in $chunkSuccess) {
                        if ($revokeOutcomes.ContainsKey($sid)) { continue }
                        $rr = if ($rBatch.Results.ContainsKey("v$sid")) { $rBatch.Results["v$sid"] } else { $null }
                        if ($null -ne $rr -and $rr.Status -ge 200 -and $rr.Status -lt 300) {
                            $revokeOutcomes[$sid] = @{ Status='Revoked'; Message='Refresh tokens and session cookies invalidated' }
                            $Counters.Revoked++
                        } else {
                            $rMsg = if ($null -eq $rr) { 'No response' } else { Get-BatchErrorText -Body $rr.Body }
                            $revokeOutcomes[$sid] = @{ Status='Failed'; Message=$rMsg }
                            $Counters.RevokeFailed++
                            if ($rMsg -match 'Authorization_RequestDenied' -or ($null -ne $rr -and $rr.Status -eq 403)) { $rAuth++ }
                        }
                    }
                    if ($rAuth -gt 0 -and $rAuth -eq $revReqs.Count) {
                        $RevokeAuthFails++
                        if ($RevokeAuthFails -ge 2) {
                            $RevokeDisabledMidRun = $true
                            Write-Host ""
                            Write-Host "[NOTICE] Force sign-out disabled - every revoke in 2 batches returned 403." -ForegroundColor Yellow
                            Write-Host "         Grant one of: $($RevokeScopes -join ', ')" -ForegroundColor DarkGray
                            Write-Host "         PASSWORD RESETS CONTINUE NORMALLY." -ForegroundColor Green
                            Write-Host ""
                        }
                    } else { $RevokeAuthFails = 0 }
                }
                catch {
                    foreach ($sid in $chunkSuccess) {
                        if (-not $revokeOutcomes.ContainsKey($sid)) {
                            $revokeOutcomes[$sid] = @{ Status='Failed'; Message="Revoke batch failed: $($_.Exception.Message)" }
                            $Counters.RevokeFailed++
                        }
                    }
                }
            }
        }
        elseif ($RevokeSessions -and $RevokeDisabledMidRun) {
            foreach ($sid in $chunkSuccess) {
                $revokeOutcomes[$sid] = @{ Status='Skipped'; Message='Force sign-out disabled mid-run' }
                $Counters.RevokeSkipped++
            }
        }

        # --- emit audit rows for the chunk, in chunk order ---
        $chunkMs = [int]((Get-Date) - $chunkStart).TotalMilliseconds
        $perRowMs = [int]($chunkMs / [Math]::Max(1, $chunk.Count))
        for ($i = 0; $i -lt $chunk.Count; $i++) {
            $sid = "r$i"
            $o   = $rowOutcomes[$sid]
            $rv  = if ($revokeOutcomes.ContainsKey($sid)) { $revokeOutcomes[$sid] } else { @{ Status=''; Message='' } }
            if ($o.Status -eq 'Success') { $Counters.Success++ } else { $Counters.Failed++ }
            $RowIndex++
            $Log.Add((New-AuditRow @AuditCommon -Row $chunk[$i] -Status $o.Status -Message $o.Message `
                        -HttpStatusCode $o.Http -GraphErrorCode $o.Code -DurationMs $perRowMs `
                        -RevokeStatus $rv.Status -RevokeMessage $rv.Message -WriteMode 'Batched'))
        }

        # --- circuit breaker: a whole chunk of 403s means the operator, not the data ---
        $chunk403 = @($rowOutcomes.Values | Where-Object { $_.Code -eq 'Authorization_RequestDenied' -or $_.Http -eq '403' }).Count
        if ($chunk403 -eq $chunk.Count) { $ConsecutiveAuthFails += $chunk403 } else { $ConsecutiveAuthFails = 0 }

        # --- governor ---
        $chunkSecs   = ((Get-Date) - $chunkStart).TotalSeconds
        $minSecs     = $chunk.Count / [double]$TargetWritesPerSec
        $throttledNow = ($null -ne $resetResults) -and (@($rowOutcomes.Values | Where-Object { $_.Http -eq '429' }).Count -gt 0)
        if ($throttledNow) {
            $GovDelayMs   = [int][Math]::Min(10000, [Math]::Max(500, $GovDelayMs * 2))
            $GovCleanRuns = 0
            Write-Host ("    [GOVERNOR] 429 seen - inter-batch delay raised to {0}ms" -f $GovDelayMs) -ForegroundColor DarkYellow
        }
        else {
            $GovCleanRuns++
            if ($GovCleanRuns -ge 5 -and $GovDelayMs -gt 0) {
                $GovDelayMs   = [int]($GovDelayMs / 2)
                $GovCleanRuns = 0
                Write-Host ("    [GOVERNOR] steady - inter-batch delay lowered to {0}ms" -f $GovDelayMs) -ForegroundColor DarkGray
            }
        }
        # Proactive pacing: never exceed the target write rate even when Graph is fast.
        $sleepSecs = [Math]::Max(($minSecs - $chunkSecs), ($GovDelayMs / 1000.0))
        if ($sleepSecs -gt 0) { Start-Sleep -Milliseconds ([int]($sleepSecs * 1000)) }

        # --- progress ---
        $okSoFar = $Counters.Success
        $elapsed = ((Get-Date) - $RunStart).TotalSeconds
        $rate    = if ($elapsed -gt 0) { $RowIndex / $elapsed } else { 0 }
        $line    = "  [BATCH {0}/{1}] rows {2}/{3} ({4:N1}%) - OK {5} / FAIL {6}" -f `
                   $ChunkIndex, $ChunkCount, $RowIndex, $Total,
                   (100.0 * $RowIndex / [Math]::Max(1, $Total)), $okSoFar, $Counters.Failed
        if ($RevokeSessions) { $line += " / OUT $($Counters.Revoked)" }
        if ($rate -gt 0) {
            $line += " | {0:N1} rows/s | ETA {1}" -f $rate, (Format-Duration (($Total - $RowIndex) / $rate))
        }
        Write-Host $line -ForegroundColor Magenta

        $FlushedRows = Save-LogIncremental -Log $Log -Path $LogPath -AlreadyFlushed $FlushedRows

        if ($ConsecutiveAuthFails -ge $AuthFailThreshold) {
            $AbortReason = "$ConsecutiveAuthFails consecutive Authorization_RequestDenied (403) failures"
            Write-Host ""
            Write-Host "============================================================" -ForegroundColor Red
            Write-Host " ABORTED BY CIRCUIT BREAKER" -ForegroundColor Red
            Write-Host "============================================================" -ForegroundColor Red
            Write-Host " $AbortReason." -ForegroundColor Red
            Write-Host " Stopped at batch $ChunkIndex of $ChunkCount. $($WorkRows.Count - $RowIndex + $Counters.Skipped) reset(s) NOT attempted." -ForegroundColor Yellow
            Write-Host " Scope? Consent? Role? PIM activation? See the guidance in Phase 2." -ForegroundColor Cyan
            Write-Host "============================================================" -ForegroundColor Red
            break
        }
    }

    if ($GovThrottled -gt 0) {
        Write-Host ""
        Write-Host "[INFO] $GovThrottled batch(es) hit throttling and were retried by the governor." -ForegroundColor DarkYellow
    }
}

# ============================================================
# PHASE 5B - SEQUENTIAL WRITE PATH
# ============================================================
foreach ($row in $(if ($UseBatchWrites) { @() } else { $ProcessedRows })) {
    $RowIndex++
    $rowStart = Get-Date
    $status    = 'Failed'
    $message   = ''
    $graphErr  = ''
    $httpCode  = ''
    $graphReqId = ''
    $attempts  = 1
    $revokeStatus  = ''
    $revokeMessage = ''

    if ($row.PreflightStatus -ne 'Ready') {
        $status  = 'Skipped'
        $message = $row.PreflightMessage
        Write-Host "[Row $RowIndex/$Total] [SKIP] $($row.UpnClean) - $message" -ForegroundColor Yellow
        $Counters.Skipped++
    }
    else {
        try {
            # Masked on purpose: never print a live credential to the console.
            Write-Host "[Row $RowIndex/$Total] Reset $($row.UpnClean) [$($row.PasswordSource)] $(Format-MaskedSecret $row.Password)" -ForegroundColor Cyan

            $PasswordProfile = @{
                Password                      = $row.Password
                ForceChangePasswordNextSignIn = $true
            }
            $attemptRef = [ref]1
            Invoke-GraphWithRetry -AttemptCounter $attemptRef -ScriptBlock {
                Update-MgUser -UserId $row.UpnClean -PasswordProfile $PasswordProfile -ErrorAction Stop
            } | Out-Null
            $attempts = $attemptRef.Value
            if ($attempts -gt 1) { $Counters.Retried++ }

            $status  = 'Success'
            $message = 'Password reset successful'
            $Counters.Success++
            $ConsecutiveAuthFails = 0
            Write-Host "[Row $RowIndex/$Total] [OK]   $($row.UpnClean)" -ForegroundColor Green

            # ---- Force sign-out (only after the password actually changed) ----
            if ($RevokeSessions -and -not $RevokeDisabledMidRun) {
                if ($row.UpnClean -eq $OperatorAccount) {
                    # Revoking our own session would kill the token running this loop.
                    $revokeStatus  = 'Skipped'
                    $revokeMessage = 'Operator account - sign-out skipped to keep this run alive'
                    $Counters.RevokeSkipped++
                    Write-Host "[Row $RowIndex/$Total] [SKIP] sign-out for your own account (would end this run)" -ForegroundColor Yellow
                }
                else {
                    try {
                        Invoke-GraphWithRetry -ScriptBlock {
                            Revoke-MgUserSignInSession -UserId $row.UpnClean -ErrorAction Stop
                        } | Out-Null
                        $revokeStatus  = 'Revoked'
                        $revokeMessage = 'Refresh tokens and session cookies invalidated'
                        $Counters.Revoked++
                        $RevokeAuthFails = 0
                        Write-Host "[Row $RowIndex/$Total] [OUT]  signed out everywhere" -ForegroundColor DarkGreen
                    }
                    catch {
                        $rRaw = [string]$_.Exception.Message
                        $revokeStatus  = 'Failed'
                        $revokeMessage = (@($rRaw -split "`r?`n") | Where-Object { $_.Trim() } | Select-Object -First 1).Trim()
                        $Counters.RevokeFailed++
                        Write-Host "[Row $RowIndex/$Total] [WARN] password WAS reset but sign-out failed - $revokeMessage" -ForegroundColor DarkYellow

                        if ($rRaw -match 'Authorization_RequestDenied' -or $rRaw -match 'Status\s*:\s*403') {
                            $RevokeAuthFails++
                            if ($RevokeAuthFails -ge 3) {
                                $RevokeDisabledMidRun = $true
                                Write-Host ""
                                Write-Host "[NOTICE] Force sign-out disabled after 3 consecutive 403s." -ForegroundColor Yellow
                                Write-Host "         The token lacks a revoke-capable scope; grant one of:" -ForegroundColor Yellow
                                Write-Host "         $($RevokeScopes -join ', ')" -ForegroundColor DarkGray
                                Write-Host "         PASSWORD RESETS CONTINUE NORMALLY - only sign-out is off." -ForegroundColor Green
                                Write-Host ""
                            }
                        } else {
                            $RevokeAuthFails = 0
                        }
                    }
                }
            }
            elseif ($RevokeSessions -and $RevokeDisabledMidRun) {
                $revokeStatus  = 'Skipped'
                $revokeMessage = 'Force sign-out disabled mid-run after repeated 403s'
                $Counters.RevokeSkipped++
            }
        }
        catch {
            $status = 'Failed'

            # The Graph SDK crams status line, ALL response headers and server diagnostics into
            # Exception.Message. Dumping that per row buries the run in ~25 lines of noise, so
            # parse out the parts that matter and keep the console to one line.
            $rawMsg = [string]$_.Exception.Message
            if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $rawMsg = "$rawMsg`n$($_.ErrorDetails.Message)" }

            $message = (@($rawMsg -split "`r?`n") | Where-Object { $_.Trim() } | Select-Object -First 1).Trim()

            if     ($rawMsg -match 'ErrorCode\s*:\s*(\S+)')            { $graphErr  = $Matches[1] }
            elseif ($rawMsg -match '^\s*\[([A-Za-z_]+)\]')             { $graphErr  = $Matches[1] }
            if     ($rawMsg -match 'Status\s*:\s*(\d{3})')             { $httpCode  = $Matches[1] }
            if     ($rawMsg -match '(?m)^\s*request-id\s*:\s*(\S+)')   { $graphReqId = $Matches[1] }
            if (-not $httpCode) { try { $httpCode = [string][int]$_.Exception.Response.StatusCode } catch { } }

            $Counters.Failed++
            Write-Host "[Row $RowIndex/$Total] [FAIL] $($row.UpnClean) - $message" -ForegroundColor Red
            if ($graphReqId) { Write-Host "         request-id: $graphReqId (quote this to Microsoft support)" -ForegroundColor DarkGray }

            # Track tenant-wide authorization denials for the circuit breaker
            if ($graphErr -eq 'Authorization_RequestDenied' -or $httpCode -eq '403') {
                $ConsecutiveAuthFails++
            } else {
                $ConsecutiveAuthFails = 0
            }
        }
    }

    $durationMs = [int]((Get-Date) - $rowStart).TotalMilliseconds

    # AUDIT LOG - NO CLEARTEXT PASSWORD. Source + length + run-salted fingerprint only.
    $Log.Add((New-AuditRow @AuditCommon -Row $row -Status $status -Message $message `
                -HttpStatusCode $httpCode -GraphErrorCode $graphErr -GraphRequestId $graphReqId `
                -AttemptCount $attempts -DurationMs $durationMs `
                -RevokeStatus $revokeStatus -RevokeMessage $revokeMessage -WriteMode 'Sequential'))

    if ($HeartbeatEvery -gt 0 -and ($RowIndex % $HeartbeatEvery) -eq 0) {
        $hb = "  [HEARTBEAT] Row {0}/{1} ({2:N1}%) - OK {3} / FAIL {4} / SKIP {5}" -f `
              $RowIndex, $Total, (100.0 * $RowIndex / [Math]::Max(1, $Total)),
              $Counters.Success, $Counters.Failed, $Counters.Skipped
        if ($RevokeSessions) { $hb += " / OUT {0}" -f $Counters.Revoked }

        # Live ETA from the observed rate, so a long run is predictable rather than a mystery.
        $elapsedSecs = ((Get-Date) - $RunStart).TotalSeconds
        if ($elapsedSecs -gt 0 -and $RowIndex -lt $Total) {
            $rate = $RowIndex / $elapsedSecs
            if ($rate -gt 0) {
                $hb += " | {0:N1} rows/s | elapsed {1} | ETA {2}" -f `
                       $rate, (Format-Duration $elapsedSecs), (Format-Duration (($Total - $RowIndex) / $rate))
            }
        }
        Write-Host $hb -ForegroundColor Magenta
    }

    if (($RowIndex % $FlushEvery) -eq 0) {
        $FlushedRows = Save-LogIncremental -Log $Log -Path $LogPath -AlreadyFlushed $FlushedRows
    }

    # CIRCUIT BREAKER: stop hammering Graph once it is clear the problem is the operator,
    # not the data. Nothing has been changed on the remaining users.
    if ($ConsecutiveAuthFails -ge $AuthFailThreshold) {
        $AbortReason = "$ConsecutiveAuthFails consecutive Authorization_RequestDenied (403) failures"
        $FlushedRows = Save-LogIncremental -Log $Log -Path $LogPath -AlreadyFlushed $FlushedRows
        Write-Host ""
        Write-Host "============================================================" -ForegroundColor Red
        Write-Host " ABORTED BY CIRCUIT BREAKER" -ForegroundColor Red
        Write-Host "============================================================" -ForegroundColor Red
        Write-Host " $AbortReason." -ForegroundColor Red
        Write-Host " Stopped at row $RowIndex of $Total. $($Total - $RowIndex) row(s) NOT attempted." -ForegroundColor Yellow
        Write-Host ""
        Write-Host " This is an OPERATOR permission problem, not a data problem. Check, in order:" -ForegroundColor Cyan
        Write-Host "   1. Scope   - the token needs '$CriticalScope'." -ForegroundColor White
        Write-Host "                User.ReadWrite.All alone has not been sufficient since 2024-12-23." -ForegroundColor DarkGray
        Write-Host "   2. Consent - a Global Administrator may need to grant admin consent to the" -ForegroundColor White
        Write-Host "                'Microsoft Graph Command Line Tools' enterprise application." -ForegroundColor DarkGray
        Write-Host "   3. Role    - you need User Administrator (for members) or Privileged" -ForegroundColor White
        Write-Host "                Authentication Administrator (to touch other admins)." -ForegroundColor DarkGray
        Write-Host "   4. PIM     - an ELIGIBLE role is not an ACTIVE role. Activate, then reconnect." -ForegroundColor White
        Write-Host "   5. Target  - resets on other admins are denied unless you hold PAA or GA." -ForegroundColor White
        Write-Host ""
        Write-Host " Fix the cause, then re-run in [1] CSV-only mode against the FAILED_RERUN file" -ForegroundColor Cyan
        Write-Host " so users receive the passwords already recorded in the credentials file." -ForegroundColor Cyan
        Write-Host "============================================================" -ForegroundColor Red
        break
    }
}

# In batched mode skipped rows are logged up front without advancing $RowIndex, so derive
# "not attempted" from what actually landed in the log rather than from the row cursor.
$RowsAccounted = $Counters.Success + $Counters.Failed + $Counters.Skipped
$NotAttempted  = [Math]::Max(0, $Total - $RowsAccounted)

# Final flush of the audit log
$FlushedRows = Save-LogIncremental -Log $Log -Path $LogPath -AlreadyFlushed $FlushedRows
Write-Host ""
if ($FlushedRows -eq $Log.Count) {
    Write-Host "[OK] Audit log saved (no secrets): $LogPath" -ForegroundColor Green
    Write-Host "     $FlushedRows row(s) written." -ForegroundColor DarkGray
} else {
    Write-Host "[FAIL] Audit log incomplete: $FlushedRows of $($Log.Count) row(s) written to $LogPath" -ForegroundColor Red
    Write-Host "       Check disk space and permissions on the log folder." -ForegroundColor Yellow
}

# Failed-only CSV, re-run ready.
# CRITICAL: it carries the SAME password that was already issued. Regenerating on retry would
# invalidate any credential you have already handed to the user.
$FailedUpns = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($l in $Log) { if ($l.Status -eq 'Failed') { [void]$FailedUpns.Add([string]$l.UpnCleaned) } }
$FailedRows = @($ProcessedRows | Where-Object { $FailedUpns.Contains($_.UpnClean) })

if ($FailedRows.Count -gt 0) {
    try {
        $FailedRows | Select-Object @{n='upn';e={ $_.UpnClean }}, @{n='password';e={ $_.Password }} |
            Export-Csv -LiteralPath $FailedLogPath -NoTypeInformation -Encoding UTF8 -Force -ErrorAction Stop
        [void](Set-SecretFileAcl -Path $FailedLogPath)
        Write-Host "[OK] Failed-only CSV (re-run ready, SAME passwords): $FailedLogPath" -ForegroundColor Green
        Write-Host "     Re-run this script against it in [1] CSV-only mode - do NOT regenerate." -ForegroundColor DarkYellow
    } catch {
        Write-Host "[WARN] Failed to export failed-only CSV: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

# Run manifest JSON
$RunEnd    = Get-Date
$DurationS = [int]($RunEnd - $RunStart).TotalSeconds
$GraphVer  = try { (Get-Module Microsoft.Graph.Users | Select-Object -First 1).Version.ToString() } catch { 'unknown' }

$Manifest = [ordered]@{
    ScriptVersion   = $ScriptVersion
    CorrelationId   = $CorrelationId
    StartTimeLocal  = $RunStart.ToString('yyyy-MM-dd HH:mm:ss')
    EndTimeLocal    = $RunEnd.ToString('yyyy-MM-dd HH:mm:ss')
    StartTimeUtc    = $RunStart.ToUniversalTime().ToString('o')
    EndTimeUtc      = $RunEnd.ToUniversalTime().ToString('o')
    DurationSeconds = $DurationS
    Operator        = $OperatorAccount
    TenantId        = $TenantId
    Host            = $env:COMPUTERNAME
    PSVersion       = $PSVersionTable.PSVersion.ToString()
    GraphSdkVersion = $GraphVer
    Input           = [ordered]@{
        CsvPath           = $CsvPath
        CsvSha256         = $CsvHash
        RowCount          = $Total
        HasPasswordColumn = $HasPasswordColumn
    }
    PasswordPolicy  = [ordered]@{
        Mode                 = $PasswordMode
        GeneratedCount       = $SourceCounts.Generated
        CsvSuppliedCount     = $SourceCounts.CSV
        GeneratorStyle       = if ($SourceCounts.Generated -gt 0) { $GenStyle } else { $null }
        GeneratorLength      = if ($SourceCounts.Generated -gt 0 -and $GenStyle -eq 'Random')   { $GenLength }    else { $null }
        GeneratorWordCount   = if ($SourceCounts.Generated -gt 0 -and $GenStyle -eq 'Readable') { $GenWordCount } else { $null }
        GeneratorDigitCount  = if ($SourceCounts.Generated -gt 0 -and $GenStyle -eq 'Readable') { $GenDigitCount } else { $null }
        EstimatedEntropyBits = if ($SourceCounts.Generated -gt 0) { $GenEntropy } else { $null }
        Rng                  = 'System.Security.Cryptography.RandomNumberGenerator'
        ForceChangeNextSignIn = $true
        UniquePerUser        = $true
    }
    ForceSignOut    = [ordered]@{
        Requested        = $RevokeSessions
        ScopeCapable     = $RevokeCapable
        Cmdlet           = 'Revoke-MgUserSignInSession'
        GraphAction      = 'POST /users/{id}/revokeSignInSessions'
        DisabledMidRun   = $RevokeDisabledMidRun
        Revoked          = $Counters.Revoked
        RevokeFailed     = $Counters.RevokeFailed
        RevokeSkipped    = $Counters.RevokeSkipped
        Caveat           = 'Invalidates refresh tokens and session cookies; already-issued access tokens remain valid until expiry (up to ~1h) unless the service supports Continuous Access Evaluation. Microsoft notes a few minutes of propagation delay.'
    }
    Outputs         = [ordered]@{
        AuditLog        = $LogPath
        CredentialsFile = $CredentialsPath
        FailedRerunFile = if ($FailedRows.Count -gt 0) { $FailedLogPath } else { $null }
        Manifest        = $ManifestPath
        SecretsRedacted = $true
        FingerprintSalt = 'CorrelationId (run-scoped; fingerprints are not comparable across runs)'
    }
    Execution       = [ordered]@{
        Aborted            = [bool]$AbortReason
        AbortReason        = if ($AbortReason) { $AbortReason } else { $null }
        RowsAttempted      = $RowsAccounted
        RowsNotAttempted   = $NotAttempted
        WriteMode          = if ($UseBatchWrites) { 'Batched' } else { 'Sequential' }
        BatchSize          = if ($UseBatchWrites) { $GraphBatchMax } else { $null }
        TargetWritesPerSec = if ($UseBatchWrites) { $TargetWritesPerSec } else { $null }
        AuditLogOrder      = if ($UseBatchWrites) { 'Execution order - sort by Row for CSV order' } else { 'CSV row order' }
        CsvRowRangeFrom    = $rangeFrom
        CsvRowRangeTo      = $rangeTo
        CsvTotalRows       = $TotalCsvRows
        ResumeLog          = if ($ResumeLogPath) { $ResumeLogPath } else { $null }
        ResumeSkipped      = $PreflightCounts.ResumeSkip
        HeartbeatEveryRows = $HeartbeatEvery
        PrivilegedScanPerformed = $ScanPrivileged
        AdminGateActive         = $ScanPrivileged
        CredentialsApplied = ($Counters.Success -gt 0)
        CredentialsFileValidFor = "$($Counters.Success) of $(@($CredRows).Count) recorded rows - do NOT distribute entries whose Status is not Success"
    }
    Counts          = [ordered]@{
        Total             = $Total
        Success           = $Counters.Success
        Failed            = $Counters.Failed
        Skipped           = $Counters.Skipped
        NotAttempted      = $NotAttempted
        Retried           = $Counters.Retried
        SessionsRevoked   = $Counters.Revoked
        RevokeFailed      = $Counters.RevokeFailed
        RevokeSkipped     = $Counters.RevokeSkipped
        PreflightHybrid   = $PreflightCounts.Hybrid
        PreflightNotFound = $PreflightCounts.NotFound
        PreflightInvalid  = $PreflightCounts.InvalidUpn
        PreflightBadPwd   = $PreflightCounts.BadPassword
        PreflightPriv     = $PreflightCounts.Privileged
        PreflightGuest    = $PreflightCounts.Guest
        PreflightDisabled = $PreflightCounts.Disabled
        DuplicateSkipped  = $PreflightCounts.DuplicateSkip
        ResumeSkipped     = $PreflightCounts.ResumeSkip
        Sanitized         = $SanitizedCount
    }
}
try {
    $Manifest | ConvertTo-Json -Depth 5 | Out-File -LiteralPath $ManifestPath -Encoding UTF8 -Force
    Write-Host "[OK] Run manifest saved: $ManifestPath" -ForegroundColor Green
} catch {
    Write-Host "[WARN] Failed to write manifest: $($_.Exception.Message)" -ForegroundColor Yellow
}

# Summary dashboard
Write-Host ""
Write-Host "=== EXECUTION SUMMARY ===" -ForegroundColor Cyan
Write-Host (" Correlation ID  : {0}" -f $CorrelationId) -ForegroundColor Magenta
Write-Host (" Total rows      : {0}" -f $Total)
Write-Host (" Succeeded       : {0}" -f $Counters.Success) -ForegroundColor Green
Write-Host (" Failed          : {0}" -f $Counters.Failed)  -ForegroundColor Red
Write-Host (" Skipped         : {0}" -f $Counters.Skipped) -ForegroundColor Yellow
Write-Host (" Not attempted   : {0}" -f $NotAttempted)     -ForegroundColor Yellow
Write-Host (" Retried         : {0}" -f $Counters.Retried) -ForegroundColor DarkYellow
if ($AbortReason) {
    Write-Host (" ABORTED         : {0}" -f $AbortReason)  -ForegroundColor Red
}
Write-Host (" Password source : {0} CSV / {1} generated ({2})" -f $SourceCounts.CSV, $SourceCounts.Generated, $PasswordMode)
if ($RevokeSessions) {
    Write-Host (" Signed out      : {0} revoked / {1} failed / {2} skipped" -f `
        $Counters.Revoked, $Counters.RevokeFailed, $Counters.RevokeSkipped) -ForegroundColor DarkGreen
    if ($RevokeDisabledMidRun) {
        Write-Host "                   (disabled mid-run after repeated 403s - resets were unaffected)" -ForegroundColor DarkYellow
    }
    if ($Counters.Revoked -gt 0) {
        Write-Host "                   Access tokens already issued stay valid up to ~1h." -ForegroundColor DarkGray
    }
} else {
    Write-Host " Signed out      : not requested" -ForegroundColor DarkGray
}
Write-Host (" Write mode      : {0}" -f $(if ($UseBatchWrites) { "Batched ($GraphBatchMax/call)" } else { 'Sequential' }))
Write-Host (" CSV rows        : {0}-{1} of {2}" -f $rangeFrom, $rangeTo, $TotalCsvRows)
if ($PreflightCounts.ResumeSkip -gt 0) {
    Write-Host (" Resumed         : {0} row(s) skipped as already done" -f $PreflightCounts.ResumeSkip) -ForegroundColor DarkGreen
}
Write-Host (" Duration        : {0}" -f (Format-Duration $DurationS))
if ($DurationS -gt 0 -and $RowsAccounted -gt 0) {
    Write-Host (" Throughput      : {0:N1} rows/sec" -f ($RowsAccounted / $DurationS)) -ForegroundColor DarkGray
}
Write-Host (" Audit log       : {0}" -f $LogPath)
Write-Host (" Credentials     : {0}" -f $CredentialsPath) -ForegroundColor Yellow
if ($FailedRows.Count -gt 0) {
    Write-Host (" Failed re-run   : {0}" -f $FailedLogPath) -ForegroundColor Yellow
}
Write-Host "=========================" -ForegroundColor Cyan
Write-Host ""

# The credentials file is pre-written, so it lists passwords for rows that may have FAILED.
# Distributing those would hand users a password their account never received.
if ($Counters.Success -eq 0) {
    Write-Host "[DO NOT DISTRIBUTE] Zero resets succeeded. The credentials file lists passwords" -ForegroundColor Red
    Write-Host "                    that were NEVER applied to any account. Delete it, fix the" -ForegroundColor Red
    Write-Host "                    cause above, and re-run." -ForegroundColor Red
}
elseif ($Counters.Failed -gt 0 -or $NotAttempted -gt 0) {
    Write-Host "[PARTIAL RUN] The credentials file lists $(@($CredRows).Count) row(s) but only $($Counters.Success) reset(s) succeeded." -ForegroundColor Yellow
    Write-Host "              Distribute ONLY the users whose Status is 'Success' in the audit log:" -ForegroundColor Yellow
    Write-Host "                Import-Csv '$LogPath' | Where-Object Status -eq 'Success' | Select-Object UpnCleaned" -ForegroundColor White
    Write-Host "              Then delete the credentials file." -ForegroundColor Yellow
}
else {
    Write-Host "[ACTION REQUIRED] Distribute the credentials file, then delete it." -ForegroundColor Yellow
}

# Exact restart instructions when anything was left undone.
if ($NotAttempted -gt 0 -or $Counters.Failed -gt 0) {
    Write-Host ""
    Write-Host "[TO FINISH THIS JOB] Re-run the script and, at the prompts:" -ForegroundColor Cyan
    Write-Host "  Row range     : $rangeFrom to $rangeTo   (same slice)" -ForegroundColor White
    Write-Host "  Resume log    : $LogPath" -ForegroundColor White
    Write-Host "  Password mode : [1] CSV only, pointed at the credentials file, so users keep" -ForegroundColor White
    Write-Host "                  the passwords already recorded for them." -ForegroundColor DarkGray
    Write-Host "  The $($Counters.Success) user(s) already done will be skipped automatically." -ForegroundColor DarkGray
}

# Disconnect
try {
    Disconnect-MgGraph -ErrorAction Stop | Out-Null
    Write-Host "Disconnected from Microsoft Graph." -ForegroundColor DarkGray
}
catch {
    Write-Host "[WARN] Disconnect-MgGraph failed: $($_.Exception.Message)" -ForegroundColor Yellow
}
