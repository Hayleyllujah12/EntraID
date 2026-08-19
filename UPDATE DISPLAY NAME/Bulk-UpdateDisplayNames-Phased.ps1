<#
============================================================
Bulk Update Entra ID Display Names - Phased Execution
============================================================
Description:
- Designed for copy-paste execution into a PowerShell terminal
- Runs in 5 distinct phases with user confirmation
- Updates ONLY the DisplayName attribute on existing Entra ID users
- Matches users by UserPrincipalName (UPN)
- Logs all results to a timestamped CSV file
- Includes sanitization, retry/backoff, and progress checkpoints

Expected CSV columns (case-sensitive headers):
  UserPrincipalName, DisplayName

Optional headers (ignored if present): GivenName, Surname, etc.

Requirements:
- PowerShell 7+
- Microsoft.Graph.Users module
- Admin role: User Administrator or Global Administrator
- Scope: User.ReadWrite.All

How to use:
- Copy and paste each phase one at a time into your PowerShell terminal
- OR copy and paste the entire script at once
============================================================
#>


# ============================================================
# PHASE 1 - Configure Paths (Input CSV + Output Log Folder)
# ============================================================
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " PHASE 1: Configure Paths" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

# Default paths (edit these or you'll be prompted to confirm/change them)
$DefaultCsvPath   = "C:\Users\hayleyllujah\Downloads\UpdateDisplayNames.csv"
$DefaultLogFolder = "C:\Users\hayleyllujah\Downloads\Logs"

# Helper: strip invisible Unicode formatting chars (LRM/RLM/LRE/RLE/PDF/LRO/RLO/ZWSP/ZWNJ/ZWJ/BOM)
# that Windows Explorer's "Copy as path" sometimes injects.
function Remove-InvisibleChars {
    param([string]$s)
    if ([string]::IsNullOrEmpty($s)) { return $s }
    return ($s -replace '[\u200B-\u200F\u202A-\u202E\u2066-\u2069\uFEFF]', '')
}

# Prompt for CSV path (press Enter to accept default)
Write-Host ""
Write-Host "Default CSV path:" -ForegroundColor Gray
Write-Host "  $DefaultCsvPath" -ForegroundColor DarkGray
$InputCsv = Read-Host "Enter CSV file path (or press Enter to use default)"
if ([string]::IsNullOrWhiteSpace($InputCsv)) {
    $CsvPath = $DefaultCsvPath
} else {
    $CsvPath = (Remove-InvisibleChars $InputCsv).Trim('"').Trim("'").Trim()
}

# Prompt for log folder
Write-Host ""
Write-Host "Default log folder:" -ForegroundColor Gray
Write-Host "  $DefaultLogFolder" -ForegroundColor DarkGray
$InputLogFolder = Read-Host "Enter log output folder (or press Enter to use default)"
if ([string]::IsNullOrWhiteSpace($InputLogFolder)) {
    $LogFolder = $DefaultLogFolder
} else {
    $LogFolder = (Remove-InvisibleChars $InputLogFolder).Trim('"').Trim("'").Trim()
}

# Validate CSV exists
if (-not (Test-Path -LiteralPath $CsvPath)) {
    Write-Host ""
    Write-Host "ERROR: CSV file not found at:" -ForegroundColor Red
    Write-Host "  $CsvPath" -ForegroundColor Red
    $codes = ([char[]]$CsvPath | Select-Object -First 5 | ForEach-Object { [int]$_ }) -join ','
    Write-Host "  First 5 char codes: $codes  (normal 'C:\' starts with 67,58,92)" -ForegroundColor DarkGray
    Write-Host "Please correct the path and run Phase 1 again." -ForegroundColor Yellow
    return
}

# Create log folder if missing
if (-not (Test-Path -LiteralPath $LogFolder)) {
    try {
        New-Item -Path $LogFolder -ItemType Directory -Force | Out-Null
        Write-Host "Created log folder: $LogFolder" -ForegroundColor DarkGray
    }
    catch {
        Write-Host "ERROR: Could not create log folder '$LogFolder': $($_.Exception.Message)" -ForegroundColor Red
        return
    }
}

# Build the timestamped log file path
$TimeStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$LogPath   = Join-Path -Path $LogFolder -ChildPath "UpdateDisplayNamesLog_$TimeStamp.csv"

Write-Host ""
Write-Host "Paths configured:" -ForegroundColor Green
Write-Host "  CSV Input  : $CsvPath" -ForegroundColor Green
Write-Host "  Log Output : $LogPath" -ForegroundColor Green



# ============================================================
# PHASE 2 - Connect to Microsoft Graph
# ============================================================
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " PHASE 2: Connect to Microsoft Graph" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

# Ensure Microsoft.Graph.Users module is available
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Users)) {
    Write-Host ""
    Write-Host "ERROR: Microsoft.Graph.Users module is not installed." -ForegroundColor Red
    Write-Host "Install with this command, then re-run Phase 2:" -ForegroundColor Yellow
    Write-Host "  Install-Module Microsoft.Graph -Scope CurrentUser -Force" -ForegroundColor Yellow
    return
}

try {
    Import-Module Microsoft.Graph.Users -ErrorAction Stop
}
catch {
    Write-Host "ERROR: Failed to import Microsoft.Graph.Users: $($_.Exception.Message)" -ForegroundColor Red
    return
}

# Prompt for Tenant ID
$DefaultTenantId = "9409cd94-ce07-4e48-9488-6d675889af18"

Write-Host ""
Write-Host "Default Tenant ID: $DefaultTenantId" -ForegroundColor Gray
$InputTenantId = Read-Host "Enter Tenant ID (or press Enter to use default)"
if ([string]::IsNullOrWhiteSpace($InputTenantId)) {
    $TenantId = $DefaultTenantId
} else {
    $TenantId = (Remove-InvisibleChars $InputTenantId).Trim()
}

# Connect to Graph (interactive sign-in)
try {
    Write-Host ""
    Write-Host "Connecting to Microsoft Graph (Tenant: $TenantId)..." -ForegroundColor Cyan
    Write-Host "You may be prompted to sign in." -ForegroundColor DarkGray
    Connect-MgGraph -TenantId $TenantId -Scopes "User.ReadWrite.All" -NoWelcome -ErrorAction Stop
    $Context = Get-MgContext
    Write-Host ""
    Write-Host "Connected successfully." -ForegroundColor Green
    Write-Host "  Account : $($Context.Account)" -ForegroundColor Green
    Write-Host "  Tenant  : $($Context.TenantId)" -ForegroundColor Green
}
catch {
    Write-Host "ERROR: Failed to connect to Microsoft Graph: $($_.Exception.Message)" -ForegroundColor Red
    return
}



# ============================================================
# PHASE 3 - Verify CSV Contents (Read and Display)
# ============================================================
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " PHASE 3: Verify CSV Contents" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

try {
    $Users = Import-Csv -Path $CsvPath -ErrorAction Stop
}
catch {
    Write-Host "ERROR: Failed to read CSV: $($_.Exception.Message)" -ForegroundColor Red
    return
}

if ($null -eq $Users -or $Users.Count -eq 0) {
    Write-Host "ERROR: CSV file is empty: $CsvPath" -ForegroundColor Red
    return
}

# Required columns
$RequiredColumns = @('UserPrincipalName','DisplayName')
$ActualColumns   = $Users[0].PSObject.Properties.Name
$MissingColumns  = $RequiredColumns | Where-Object { $_ -notin $ActualColumns }

if ($MissingColumns.Count -gt 0) {
    Write-Host "ERROR: CSV is missing required column(s): $($MissingColumns -join ', ')" -ForegroundColor Red
    Write-Host "Required headers : $($RequiredColumns -join ', ')" -ForegroundColor Yellow
    return
}

# Display the contents in a table for review
Write-Host ""
Write-Host "CSV loaded successfully. Total rows: $($Users.Count)" -ForegroundColor Green
Write-Host ""
Write-Host "Preview of CSV contents (first 20 rows):" -ForegroundColor Cyan
Write-Host "------------------------------------------------------------" -ForegroundColor DarkGray

$RowNum = 0
$Users | Select-Object -First 20 | ForEach-Object {
    $RowNum++
    [PSCustomObject]@{
        Row             = $RowNum
        UPN             = $(if ([string]::IsNullOrWhiteSpace($_.UserPrincipalName)) { '<EMPTY>' } else { $_.UserPrincipalName.Trim() })
        NewDisplayName  = $(if ([string]::IsNullOrWhiteSpace($_.DisplayName))       { '<EMPTY>' } else { $_.DisplayName.Trim() })
    }
} | Format-Table -AutoSize | Out-Host

if ($Users.Count -gt 20) {
    Write-Host "  ... and $($Users.Count - 20) more row(s) not shown" -ForegroundColor DarkGray
}
Write-Host "------------------------------------------------------------" -ForegroundColor DarkGray



# ============================================================
# PHASE 4 - Confirm Before Proceeding (Yes / No)
# ============================================================
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " PHASE 4: Confirmation" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "You are about to UPDATE the DisplayName for up to $($Users.Count) user(s)" -ForegroundColor Yellow
Write-Host "in tenant:" -ForegroundColor Yellow
Write-Host "  $($Context.TenantId)" -ForegroundColor Yellow
Write-Host ""
Write-Host "ONLY the DisplayName attribute will be modified." -ForegroundColor Yellow
Write-Host "No other attributes (licenses, password, department, etc.) will be touched." -ForegroundColor Yellow
Write-Host ""
Write-Host "Users not found in Entra ID will be logged as 'Skipped'." -ForegroundColor Yellow
Write-Host "Users whose DisplayName already matches the CSV value will be logged as 'NoChange'." -ForegroundColor Yellow
Write-Host ""

$Confirm = Read-Host "Is the CSV data correct and do you want to proceed? (Y/N)"

if ($Confirm -notmatch '^(y|yes)$') {
    Write-Host ""
    Write-Host "Operation cancelled by user. No changes were made." -ForegroundColor Red
    Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor DarkGray
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    return
}

Write-Host ""
Write-Host "Confirmed. Proceeding with DisplayName updates..." -ForegroundColor Green



# ============================================================
# PHASE 5 - Update DisplayName, Export Log
# ============================================================
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " PHASE 5: Update DisplayName" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

$Log = [System.Collections.Generic.List[object]]::new()
$Counters = @{ Updated = 0; NoChange = 0; Skipped = 0; Failed = 0 }

# Helper: sanitize DisplayName - strips invisible/control chars, normalizes whitespace, enforces length
function Sanitize-DisplayName {
    param([string]$Value)
    if ([string]::IsNullOrEmpty($Value)) { return @{ Value = ''; Changed = $false; Reason = '' } }
    $original = $Value
    # Strip invisible formatting chars
    $cleaned = $Value -replace '[\u200B-\u200F\u202A-\u202E\u2066-\u2069\uFEFF]', ''
    # Strip control chars (except normal whitespace)
    $cleaned = $cleaned -replace '[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F-\u009F]', ''
    # Collapse internal whitespace runs to a single space and trim
    $cleaned = ($cleaned -replace '\s+', ' ').Trim()
    # Enforce DisplayName max length (256)
    $truncated = $false
    if ($cleaned.Length -gt 256) {
        $cleaned = $cleaned.Substring(0, 256).Trim()
        $truncated = $true
    }
    $changed = ($cleaned -ne $original)
    $reason = ''
    if ($changed) {
        $r = @()
        if ($original -match '[\u200B-\u200F\u202A-\u202E\u2066-\u2069\uFEFF]') { $r += 'invisible chars' }
        if ($original -match '[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F-\u009F]') { $r += 'control chars' }
        if ($truncated) { $r += 'truncated to 256' }
        if (-not $r) { $r += 'whitespace normalized' }
        $reason = "DisplayName sanitized ($($r -join ', '))"
    }
    return @{ Value = $cleaned; Changed = $changed; Reason = $reason }
}

# Helper: invoke a Graph call with retry/backoff for 429 throttling and transient 5xx errors
function Invoke-GraphWithRetry {
    param(
        [Parameter(Mandatory)][scriptblock]$Action,
        [int]$MaxAttempts = 5,
        [int]$BaseDelaySeconds = 5
    )
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            return & $Action
        }
        catch {
            $isLast = ($attempt -eq $MaxAttempts)
            $msg = $_.Exception.Message
            $status = $null
            if ($_.Exception.Response -and $_.Exception.Response.StatusCode) {
                $status = [int]$_.Exception.Response.StatusCode
            }
            $isRetryable = ($status -eq 429) -or ($status -ge 500 -and $status -lt 600) -or ($msg -match 'throttl|TooManyRequests|timed out|timeout')
            if (-not $isRetryable -or $isLast) { throw }

            $delay = $BaseDelaySeconds * [math]::Pow(2, $attempt - 1)
            if ($_.Exception.Response -and $_.Exception.Response.Headers) {
                $retryAfter = $_.Exception.Response.Headers['Retry-After']
                if ($retryAfter -and ($retryAfter -as [int])) { $delay = [int]$retryAfter }
            }
            Write-Host "    Throttled/transient error (status=$status). Retrying in $delay s... (attempt $attempt/$MaxAttempts)" -ForegroundColor DarkYellow
            Start-Sleep -Seconds $delay
        }
    }
}

$ProgressEvery = 25
$TotalUsers    = $Users.Count
$RowNumber     = 0

foreach ($u in $Users) {
    $RowNumber++

    # Heartbeat every N rows so a stuck run is obvious
    if (($RowNumber % $ProgressEvery) -eq 0) {
        Write-Host ""
        Write-Host "--- Progress: $RowNumber / $TotalUsers (Updated=$($Counters.Updated), NoChange=$($Counters.NoChange), Skipped=$($Counters.Skipped), Failed=$($Counters.Failed)) ---" -ForegroundColor Magenta
        Write-Host ""
    }

    # Periodic log flush every 100 rows (so Ctrl+C doesn't lose progress)
    if (($RowNumber % 100) -eq 0 -and $Log.Count -gt 0) {
        try {
            $Log | Export-Csv -Path $LogPath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
        } catch { }
    }

    $UPN            = if ($null -ne $u.UserPrincipalName) { ([string]$u.UserPrincipalName).Trim() } else { "" }
    $NewDisplayName = if ($null -ne $u.DisplayName)       { ([string]$u.DisplayName).Trim() }       else { "" }

    # Skip empty required fields
    if ([string]::IsNullOrWhiteSpace($UPN)) {
        Write-Host "[Row $RowNumber] Skipped: UPN is empty" -ForegroundColor Yellow
        $Log.Add([PSCustomObject]@{
            Timestamp         = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
            Row               = $RowNumber
            UserPrincipalName = ""
            OldDisplayName    = ""
            NewDisplayName    = $NewDisplayName
            Status            = "Skipped"
            Message           = "UPN is empty"
        })
        $Counters.Skipped++
        continue
    }

    if ([string]::IsNullOrWhiteSpace($NewDisplayName)) {
        Write-Host "[Row $RowNumber] Skipped ${UPN}: DisplayName is empty" -ForegroundColor Yellow
        $Log.Add([PSCustomObject]@{
            Timestamp         = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
            Row               = $RowNumber
            UserPrincipalName = $UPN
            OldDisplayName    = ""
            NewDisplayName    = ""
            Status            = "Skipped"
            Message           = "DisplayName is empty"
        })
        $Counters.Skipped++
        continue
    }

    # Sanitize the new display name
    $San = Sanitize-DisplayName -Value $NewDisplayName
    if ($San.Changed) {
        Write-Host "[Row $RowNumber]   $($San.Reason)" -ForegroundColor DarkYellow
    }
    $NewDisplayName = $San.Value

    # Re-check after sanitization
    if ([string]::IsNullOrWhiteSpace($NewDisplayName)) {
        Write-Host "[Row $RowNumber] Skipped ${UPN}: DisplayName empty after sanitization" -ForegroundColor Yellow
        $Log.Add([PSCustomObject]@{
            Timestamp         = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
            Row               = $RowNumber
            UserPrincipalName = $UPN
            OldDisplayName    = ""
            NewDisplayName    = ""
            Status            = "Skipped"
            Message           = "DisplayName empty after sanitization"
        })
        $Counters.Skipped++
        continue
    }

    # Look up the user to get current DisplayName
    $ExistingUser = $null
    try {
        $ExistingUser = Invoke-GraphWithRetry {
            Get-MgUser -UserId $UPN -Property "Id,UserPrincipalName,DisplayName" -ErrorAction Stop
        }
    }
    catch {
        $errMsg = $_.Exception.Message
        if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
            $errMsg = "$errMsg | Details: $($_.ErrorDetails.Message)"
        }
        # Distinguish "not found" from other errors
        $isNotFound = $errMsg -match 'Request_ResourceNotFound|Resource ''.*'' does not exist|NotFound'
        if ($isNotFound) {
            Write-Host "[Row $RowNumber] Skipped ${UPN}: user not found in Entra" -ForegroundColor DarkYellow
            $Log.Add([PSCustomObject]@{
                Timestamp         = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
                Row               = $RowNumber
                UserPrincipalName = $UPN
                OldDisplayName    = ""
                NewDisplayName    = $NewDisplayName
                Status            = "Skipped"
                Message           = "User not found in Entra ID"
            })
            $Counters.Skipped++
        } else {
            Write-Host "[Row $RowNumber] FAILED to look up ${UPN}: $errMsg" -ForegroundColor Red
            $Log.Add([PSCustomObject]@{
                Timestamp         = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
                Row               = $RowNumber
                UserPrincipalName = $UPN
                OldDisplayName    = ""
                NewDisplayName    = $NewDisplayName
                Status            = "Failed"
                Message           = "Lookup failed: $errMsg"
            })
            $Counters.Failed++
        }
        continue
    }

    $OldDisplayName = if ($null -ne $ExistingUser.DisplayName) { [string]$ExistingUser.DisplayName } else { "" }

    # If unchanged, skip the write
    if ($OldDisplayName -eq $NewDisplayName) {
        Write-Host "[Row $RowNumber] No change for ${UPN}: DisplayName already '$NewDisplayName'" -ForegroundColor DarkGray
        $Log.Add([PSCustomObject]@{
            Timestamp         = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
            Row               = $RowNumber
            UserPrincipalName = $UPN
            OldDisplayName    = $OldDisplayName
            NewDisplayName    = $NewDisplayName
            Status            = "NoChange"
            Message           = "DisplayName already matches target"
        })
        $Counters.NoChange++
        continue
    }

    # Update
    try {
        Write-Host "[Row $RowNumber] Updating ${UPN}: '$OldDisplayName' -> '$NewDisplayName'" -ForegroundColor Cyan
        Invoke-GraphWithRetry {
            Update-MgUser -UserId $ExistingUser.Id -DisplayName $NewDisplayName -ErrorAction Stop
        }
        Write-Host "[Row $RowNumber]   SUCCESS" -ForegroundColor Green
        $Log.Add([PSCustomObject]@{
            Timestamp         = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
            Row               = $RowNumber
            UserPrincipalName = $UPN
            OldDisplayName    = $OldDisplayName
            NewDisplayName    = $NewDisplayName
            Status            = "Updated"
            Message           = "DisplayName updated from '$OldDisplayName' to '$NewDisplayName'"
        })
        $Counters.Updated++
    }
    catch {
        $errMsg = $_.Exception.Message
        if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
            $errMsg = "$errMsg | Details: $($_.ErrorDetails.Message)"
        }
        Write-Host "[Row $RowNumber]   FAILED: $errMsg" -ForegroundColor Red
        $Log.Add([PSCustomObject]@{
            Timestamp         = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
            Row               = $RowNumber
            UserPrincipalName = $UPN
            OldDisplayName    = $OldDisplayName
            NewDisplayName    = $NewDisplayName
            Status            = "Failed"
            Message           = "Update failed: $errMsg"
        })
        $Counters.Failed++
    }
}

# Export log to CSV
try {
    $Log | Export-Csv -Path $LogPath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
    Write-Host ""
    Write-Host "Log file saved to:" -ForegroundColor Green
    Write-Host "  $LogPath" -ForegroundColor Green
}
catch {
    Write-Host "ERROR: Failed to export log CSV: $($_.Exception.Message)" -ForegroundColor Red
}

# Summary
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " Summary" -ForegroundColor Cyan
Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
Write-Host (" Total rows         : {0}" -f $Users.Count)
Write-Host (" Updated            : {0}" -f $Counters.Updated)  -ForegroundColor Green
Write-Host (" No change needed   : {0}" -f $Counters.NoChange) -ForegroundColor DarkGray
Write-Host (" Skipped            : {0}" -f $Counters.Skipped)  -ForegroundColor Yellow
Write-Host (" Failed             : {0}" -f $Counters.Failed)   -ForegroundColor Red
Write-Host (" End Time           : {0}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
Write-Host "============================================================" -ForegroundColor Cyan

# Disconnect from Microsoft Graph
try {
    Disconnect-MgGraph -ErrorAction Stop | Out-Null
    Write-Host "Disconnected from Microsoft Graph." -ForegroundColor DarkGray
}
catch {
    Write-Host "Warning: Disconnect-MgGraph failed: $($_.Exception.Message)" -ForegroundColor Yellow
}
