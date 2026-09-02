<#
============================================================
Bulk Extract Entra ID Authentication Methods - Phased Execution
============================================================
Description:
- Designed for copy-paste execution into a PowerShell terminal
- Runs in 5 distinct phases with user confirmation
- Reads a CSV with header: upn
- Queries Microsoft Graph for registered authentication methods per user
- Detects: Password, Email, Microsoft Authenticator, SMS, Voice, FIDO2,
           Windows Hello for Business, Temporary Access Pass
- Infers Default Sign-In Method (Admin Center-like priority)
- Exports results to a timestamped CSV file

Expected CSV columns (case-sensitive headers):
  upn

Requirements:
- PowerShell 7+
- Microsoft.Graph.Users + Microsoft.Graph.Identity.SignIns modules
- Admin role: Global Reader / Authentication Admin / Global Admin
- Scopes: User.Read.All, UserAuthenticationMethod.Read.All

How to use:
- Copy and paste each phase one at a time into your PowerShell terminal
- OR copy and paste the entire script at once

.AUTHOR         Generated with Claude for Rakso CT Education IT.
.VERSION        2.0
.DATE           2026-08-19
.REQUIREMENTS   PowerShell 7+; Microsoft.Graph.Users, Microsoft.Graph.Identity.SignIns.
.PERMISSIONS    User.Read.All, UserAuthenticationMethod.Read.All (both read-only).
.SAFETY         Read-only. No user attribute is modified at any point.
.CHANGELOG      v1.0 - Initial release.
                v2.0 - Removed the hardcoded tenant GUID and the hardcoded
                       operator-specific input CSV / log folder paths. Tenant ID is
                       now prompted at runtime and GUID-validated before connecting;
                       path defaults derive from the script's own folder and stay
                       overridable at the prompt. The file now contains nothing
                       tenant- or machine-specific and is safe to commit.
============================================================
#>


# ============================================================
# PHASE 1 - Configure Paths (Input CSV + Output Log Folder)
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

$DefaultCsvPath   = Join-Path $ScriptDir 'Bulk-ExtractAuthMethod.csv'
$DefaultLogFolder = Join-Path $ScriptDir 'Auth_Method_logs'

# Helper: strip invisible Unicode formatting chars from pasted paths
function Remove-InvisibleChars {
    param([string]$s)
    if ([string]::IsNullOrEmpty($s)) { return $s }
    return ($s -replace '[\u200B-\u200F\u202A-\u202E\u2066-\u2069\uFEFF]', '')
}

# Prompt for CSV path
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

# Build the timestamped output file path
$TimeStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$OutputCsv = Join-Path -Path $LogFolder -ChildPath "AuthMethodLogs_$TimeStamp.csv"

Write-Host ""
Write-Host "Paths configured:" -ForegroundColor Green
Write-Host "  CSV Input  : $CsvPath" -ForegroundColor Green
Write-Host "  CSV Output : $OutputCsv" -ForegroundColor Green



# ============================================================
# PHASE 2 - Connect to Microsoft Graph
# ============================================================
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " PHASE 2: Connect to Microsoft Graph" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

# Ensure required Microsoft.Graph modules are available
$RequiredModules = @('Microsoft.Graph.Users', 'Microsoft.Graph.Identity.SignIns')
foreach ($Module in $RequiredModules) {
    if (-not (Get-Module -ListAvailable -Name $Module)) {
        Write-Host ""
        Write-Host "ERROR: Required module '$Module' is not installed." -ForegroundColor Red
        Write-Host "Install with this command, then re-run Phase 2:" -ForegroundColor Yellow
        Write-Host "  Install-Module Microsoft.Graph -Scope CurrentUser -Force" -ForegroundColor Yellow
        return
    }
    try {
        Import-Module $Module -ErrorAction Stop
    }
    catch {
        Write-Host "ERROR: Failed to import '$Module': $($_.Exception.Message)" -ForegroundColor Red
        return
    }
}

# --- Tenant ID ---------------------------------------------------------------
# Prompted at runtime and never stored in this file, so the script is safe to
# commit and safe to hand to another admin or another tenant.
Write-Host ""
$InputTenantId = Read-Host "Enter Tenant ID (GUID)"
$TenantId = (Remove-InvisibleChars $InputTenantId).Trim().Trim('"').Trim("'").Trim()
if ([string]::IsNullOrWhiteSpace($TenantId)) {
    Write-Host "ERROR: Tenant ID is required." -ForegroundColor Red
    return
}
if ($TenantId -notmatch '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$') {
    Write-Host "ERROR: '$TenantId' is not a valid GUID." -ForegroundColor Red
    Write-Host "  Find it in the Entra admin center > Overview > Tenant ID." -ForegroundColor Yellow
    return
}

# Sign-in mode selection
Write-Host ""
Write-Host "Sign-in mode:" -ForegroundColor Cyan
Write-Host "  [1] Browser session (device code - opens in your current browser, no popup)" -ForegroundColor White
Write-Host "  [2] Interactive popup window (default Graph SDK behavior)" -ForegroundColor White
Write-Host "  [3] Reuse existing session if already connected to this tenant" -ForegroundColor White
Write-Host ""
$SignInChoice = Read-Host "Enter choice (1, 2, or 3 - press Enter for 1)"
$SignInChoice = $SignInChoice.Trim()
if ([string]::IsNullOrWhiteSpace($SignInChoice)) { $SignInChoice = '1' }

# Check for existing session if option 3 chosen
$AlreadyConnected = $false
if ($SignInChoice -eq '3') {
    try {
        $ExistingContext = Get-MgContext -ErrorAction Stop
        if ($ExistingContext -and $ExistingContext.TenantId -eq $TenantId) {
            $AlreadyConnected = $true
            $Context = $ExistingContext
            Write-Host ""
            Write-Host "Reusing existing session." -ForegroundColor Green
            Write-Host "  Account : $($Context.Account)" -ForegroundColor Green
            Write-Host "  Tenant  : $($Context.TenantId)" -ForegroundColor Green
            Write-Host "  Scopes  : $($Context.Scopes -join ', ')" -ForegroundColor DarkGray
        } else {
            Write-Host "No matching existing session found - falling back to browser sign-in." -ForegroundColor DarkYellow
            $SignInChoice = '1'
        }
    }
    catch {
        Write-Host "No existing session - falling back to browser sign-in." -ForegroundColor DarkYellow
        $SignInChoice = '1'
    }
}

# Connect using the chosen method (skip if already connected via option 3)
if (-not $AlreadyConnected) {
    try {
        Write-Host ""
        Write-Host "Connecting to Microsoft Graph (Tenant: $TenantId)..." -ForegroundColor Cyan

        if ($SignInChoice -eq '1') {
            Write-Host "A device code will be shown below. Open the URL in your current browser and enter the code." -ForegroundColor DarkGray
            Connect-MgGraph -TenantId $TenantId -Scopes "User.Read.All","UserAuthenticationMethod.Read.All" -UseDeviceCode -NoWelcome -ErrorAction Stop
        }
        else {
            Write-Host "An interactive popup window will appear for sign-in." -ForegroundColor DarkGray
            Connect-MgGraph -TenantId $TenantId -Scopes "User.Read.All","UserAuthenticationMethod.Read.All" -NoWelcome -ErrorAction Stop
        }

        $Context = Get-MgContext
        Write-Host ""
        Write-Host "Connected successfully." -ForegroundColor Green
        Write-Host "  Account : $($Context.Account)" -ForegroundColor Green
        Write-Host "  Tenant  : $($Context.TenantId)" -ForegroundColor Green
        Write-Host "  Scopes  : $($Context.Scopes -join ', ')" -ForegroundColor DarkGray
    }
    catch {
        Write-Host "ERROR: Failed to connect to Microsoft Graph: $($_.Exception.Message)" -ForegroundColor Red
        return
    }
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

# Required column check (accept either 'upn' or 'UserPrincipalName')
$ActualColumns = $Users[0].PSObject.Properties.Name
$UpnColumn = $null
if ($ActualColumns -contains 'upn')                     { $UpnColumn = 'upn' }
elseif ($ActualColumns -contains 'UPN')                 { $UpnColumn = 'UPN' }
elseif ($ActualColumns -contains 'UserPrincipalName')   { $UpnColumn = 'UserPrincipalName' }

if (-not $UpnColumn) {
    Write-Host "ERROR: CSV must contain a 'upn' or 'UserPrincipalName' column." -ForegroundColor Red
    Write-Host "Found columns: $($ActualColumns -join ', ')" -ForegroundColor Yellow
    return
}

Write-Host ""
Write-Host "CSV loaded successfully. Total rows: $($Users.Count)" -ForegroundColor Green
Write-Host "Using column: $UpnColumn" -ForegroundColor DarkGray
Write-Host ""
Write-Host "Preview of CSV contents (first 20 rows):" -ForegroundColor Cyan
Write-Host "------------------------------------------------------------" -ForegroundColor DarkGray

$RowNum = 0
$Users | Select-Object -First 20 | ForEach-Object {
    $RowNum++
    [PSCustomObject]@{
        Row = $RowNum
        UPN = $(if ([string]::IsNullOrWhiteSpace($_.$UpnColumn)) { '<EMPTY>' } else { $_.$UpnColumn.Trim() })
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
Write-Host "You are about to QUERY authentication methods for $($Users.Count) user(s)" -ForegroundColor Yellow
Write-Host "in tenant:" -ForegroundColor Yellow
Write-Host "  $($Context.TenantId)" -ForegroundColor Yellow
Write-Host ""
Write-Host "This is a READ-ONLY operation. No user attributes will be modified." -ForegroundColor Green
Write-Host ""
Write-Host "The following methods will be detected per user:" -ForegroundColor Yellow
Write-Host "  - Password, Email OTP" -ForegroundColor Yellow
Write-Host "  - Microsoft Authenticator (push/number matching)" -ForegroundColor Yellow
Write-Host "  - SMS (mobile), Voice Call (office)" -ForegroundColor Yellow
Write-Host "  - FIDO2 Security Key" -ForegroundColor Yellow
Write-Host "  - Windows Hello for Business" -ForegroundColor Yellow
Write-Host "  - Temporary Access Pass" -ForegroundColor Yellow
Write-Host ""
Write-Host "Plus an inferred Default Sign-In Method (Admin Center-like priority)." -ForegroundColor Yellow
Write-Host ""

$Confirm = Read-Host "Is the CSV data correct and do you want to proceed? (Y/N)"

if ($Confirm -notmatch '^(y|yes)$') {
    Write-Host ""
    Write-Host "Operation cancelled by user. No queries were made." -ForegroundColor Red
    Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor DarkGray
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    return
}

Write-Host ""
Write-Host "Confirmed. Proceeding with authentication method extraction..." -ForegroundColor Green



# ============================================================
# PHASE 5 - Extract Auth Methods, Export Results
# ============================================================
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " PHASE 5: Extract Authentication Methods" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

$Results = [System.Collections.Generic.List[object]]::new()
$Counters = @{ Success = 0; NotFound = 0; Failed = 0; Skipped = 0 }

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

    # Heartbeat every N rows
    if (($RowNumber % $ProgressEvery) -eq 0) {
        Write-Host ""
        Write-Host "--- Progress: $RowNumber / $TotalUsers (Success=$($Counters.Success), NotFound=$($Counters.NotFound), Failed=$($Counters.Failed), Skipped=$($Counters.Skipped)) ---" -ForegroundColor Magenta
        Write-Host ""
    }

    # Periodic CSV flush every 100 rows so Ctrl+C doesn't lose progress
    if (($RowNumber % 100) -eq 0 -and $Results.Count -gt 0) {
        try {
            $Results | Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
        } catch { }
    }

    $UPN = if ($null -ne $u.$UpnColumn) { ([string]$u.$UpnColumn).Trim() } else { "" }

    if ([string]::IsNullOrWhiteSpace($UPN)) {
        Write-Host "[Row $RowNumber] Skipped: UPN is empty" -ForegroundColor Yellow
        $Results.Add([PSCustomObject]@{
            Timestamp           = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
            Row                 = $RowNumber
            UPN                 = ""
            Password            = $false
            Email               = $false
            AuthenticatorApp    = $false
            SMS                 = $false
            VoiceCall           = $false
            FIDO2               = $false
            WindowsHello        = $false
            TemporaryAccessPass = $false
            MethodCount         = 0
            DefaultSignInMethod = "N/A"
            Status              = "Skipped"
            Message             = "UPN is empty"
        })
        $Counters.Skipped++
        continue
    }

    Write-Host "[Row $RowNumber] Querying $UPN..." -ForegroundColor Cyan

    # Query Graph for auth methods with retry
    $Methods = $null
    try {
        $Methods = Invoke-GraphWithRetry {
            Get-MgUserAuthenticationMethod -UserId $UPN -ErrorAction Stop
        }
    }
    catch {
        $errMsg = $_.Exception.Message
        if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
            $errMsg = "$errMsg | Details: $($_.ErrorDetails.Message)"
        }
        $isNotFound = $errMsg -match 'Request_ResourceNotFound|Resource ''.*'' does not exist|NotFound|User Not Found'

        if ($isNotFound) {
            Write-Host "[Row $RowNumber]   User not found in Entra" -ForegroundColor DarkYellow
            $Counters.NotFound++
            $Status = "NotFound"
            $Message = "User not found in Entra ID"
        } else {
            Write-Host "[Row $RowNumber]   FAILED: $errMsg" -ForegroundColor Red
            $Counters.Failed++
            $Status = "Failed"
            $Message = $errMsg
        }

        $Results.Add([PSCustomObject]@{
            Timestamp           = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
            Row                 = $RowNumber
            UPN                 = $UPN
            Password            = $false
            Email               = $false
            AuthenticatorApp    = $false
            SMS                 = $false
            VoiceCall           = $false
            FIDO2               = $false
            WindowsHello        = $false
            TemporaryAccessPass = $false
            MethodCount         = 0
            DefaultSignInMethod = "N/A"
            Status              = $Status
            Message             = $Message
        })
        continue
    }

    # Initialize method flags
    $Password            = $false
    $Email               = $false
    $Authenticator       = $false
    $SMS                 = $false
    $VoiceCall           = $false
    $FIDO2               = $false
    $WindowsHello        = $false
    $TemporaryAccessPass = $false

    foreach ($m in $Methods) {
        switch ($m.AdditionalProperties.'@odata.type') {
            "#microsoft.graph.passwordAuthenticationMethod"               { $Password = $true }
            "#microsoft.graph.emailAuthenticationMethod"                  { $Email = $true }
            "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod" { $Authenticator = $true }
            "#microsoft.graph.phoneAuthenticationMethod" {
                if ($m.AdditionalProperties.phoneType -eq "mobile") { $SMS = $true }
                if ($m.AdditionalProperties.phoneType -eq "office") { $VoiceCall = $true }
            }
            "#microsoft.graph.fido2AuthenticationMethod"                  { $FIDO2 = $true }
            "#microsoft.graph.windowsHelloForBusinessAuthenticationMethod"{ $WindowsHello = $true }
            "#microsoft.graph.temporaryAccessPassAuthenticationMethod"    { $TemporaryAccessPass = $true }
        }
    }

    # Infer Default Sign-In Method (Admin Center-like priority)
    $DefaultSignInMethod =
        if ($FIDO2)            { "FIDO2 Security Key" }
        elseif ($WindowsHello) { "Windows Hello for Business" }
        elseif ($Authenticator){ "Microsoft Authenticator (Push / Number Matching)" }
        elseif ($Email)        { "Email OTP" }
        elseif ($SMS)          { "SMS (Primary Mobile)" }
        elseif ($VoiceCall)    { "Voice Call (Office)" }
        elseif ($Password)     { "Password" }
        else                   { "None / Unknown" }

    $MethodCount = @($Password, $Email, $Authenticator, $SMS, $VoiceCall, $FIDO2, $WindowsHello, $TemporaryAccessPass | Where-Object { $_ }).Count

    Write-Host "[Row $RowNumber]   $MethodCount method(s) registered. Default: $DefaultSignInMethod" -ForegroundColor Green

    $Results.Add([PSCustomObject]@{
        Timestamp           = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        Row                 = $RowNumber
        UPN                 = $UPN
        Password            = $Password
        Email               = $Email
        AuthenticatorApp    = $Authenticator
        SMS                 = $SMS
        VoiceCall           = $VoiceCall
        FIDO2               = $FIDO2
        WindowsHello        = $WindowsHello
        TemporaryAccessPass = $TemporaryAccessPass
        MethodCount         = $MethodCount
        DefaultSignInMethod = $DefaultSignInMethod
        Status              = "Success"
        Message             = ""
    })
    $Counters.Success++
}

# Export results to CSV
try {
    $Results | Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
    Write-Host ""
    Write-Host "Output file saved to:" -ForegroundColor Green
    Write-Host "  $OutputCsv" -ForegroundColor Green
}
catch {
    Write-Host "ERROR: Failed to export output CSV: $($_.Exception.Message)" -ForegroundColor Red
}

# Summary
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " Summary" -ForegroundColor Cyan
Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
Write-Host (" Total rows         : {0}" -f $Users.Count)
Write-Host (" Success            : {0}" -f $Counters.Success)  -ForegroundColor Green
Write-Host (" Not Found          : {0}" -f $Counters.NotFound) -ForegroundColor DarkYellow
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
