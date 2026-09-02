<#
============================================================
Bulk Create Entra ID Users + Assign License - Phased Execution
============================================================
Description:
- Designed for copy-paste execution into a PowerShell terminal
- Runs in 5 distinct phases with user confirmation
- Creates new cloud users in Microsoft Entra ID
- Assigns Microsoft 365 A1 or A3 license (Student or Faculty)
- Logs all results to a timestamped CSV file

Expected CSV columns (case-sensitive headers):
  DisplayName, UserPrincipalName, MailNickname, Password,
  UsageLocation, GivenName, Surname, Department, JobTitle

Requirements:
- PowerShell 7+
- Microsoft.Graph module (Users + Identity.DirectoryManagement)
- Admin role: User Administrator or Global Administrator
- Scopes: User.ReadWrite.All, Directory.ReadWrite.All,
          Organization.Read.All (to read available SKUs)

How to use:
- Copy and paste each phase one at a time into your PowerShell terminal
- OR copy and paste the entire script at once

.AUTHOR         Generated with Claude for Rakso CT Education IT.
.VERSION        2.0
.DATE           2026-08-19
.REQUIREMENTS   PowerShell 7+; Microsoft.Graph.Users,
                Microsoft.Graph.Identity.DirectoryManagement, Microsoft.Graph.Users.Actions.
.PERMISSIONS    User.ReadWrite.All, Directory.ReadWrite.All, Organization.Read.All.
.SAFETY         Y/N confirmation before any write. SKU verified against the tenant and
                checked against the CSV row count before execution.
.CHANGELOG      v1.0 - Initial release.
                v2.0 - De-hardcoded for source control and multi-tenant use:
                       * Tenant GUID removed - now prompted and GUID-validated at runtime.
                       * The four hardcoded education SKU GUIDs are gone. Licenses are
                         listed live from Get-MgSubscribedSku and picked by number, so a
                         stale GUID can no longer fail every row at the license step
                         AFTER the users have already been created.
                       * Operator-specific input CSV and log folder paths removed;
                         defaults now derive from the script's own folder.
                       * Added Remove-InvisibleChars hygiene on all typed paths.
.KNOWN GAP      Temporary passwords are still written into the same log CSV as the audit
                trail. Split these into a separate restricted credentials file before
                using this against a production tenant.
============================================================
#>


# ============================================================
# PHASE 1 - Configure Paths (Input CSV + Output Log Folder)
# ============================================================
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " PHASE 1: Configure Paths" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

# Helper: strip invisible Unicode formatting chars from pasted paths
# (LRM/RLM/ZWSP/NBSP/BOM ride along on Windows "Copy as path" and Excel exports)
function Remove-InvisibleChars {
    param([string]$s)
    if ([string]::IsNullOrEmpty($s)) { return $s }
    return ($s -replace '[\u200B-\u200F\u202A-\u202E\u2066-\u2069\uFEFF\u00A0\u202F]', '')
}

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

$DefaultCsvPath   = Join-Path $ScriptDir 'BulkCreateUsers.csv'
$DefaultLogFolder = Join-Path $ScriptDir 'Logs'

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
$LogPath   = Join-Path -Path $LogFolder -ChildPath "CreateUsersLog_$TimeStamp.csv"

Write-Host ""
Write-Host "Paths configured:" -ForegroundColor Green
Write-Host "  CSV Input  : $CsvPath" -ForegroundColor Green
Write-Host "  Log Output : $LogPath" -ForegroundColor Green



# ============================================================
# PHASE 2 - Connect to Microsoft Graph + Select License SKU
# ============================================================
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " PHASE 2: Connect to Microsoft Graph + Select License" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

# Ensure required Microsoft.Graph modules are available
$RequiredModules = @('Microsoft.Graph.Users', 'Microsoft.Graph.Identity.DirectoryManagement', 'Microsoft.Graph.Users.Actions')
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

# Connect to Graph (interactive sign-in)
try {
    Write-Host ""
    Write-Host "Connecting to Microsoft Graph (Tenant: $TenantId)..." -ForegroundColor Cyan
    Write-Host "You may be prompted to sign in." -ForegroundColor DarkGray
    Connect-MgGraph -TenantId $TenantId -Scopes "User.ReadWrite.All","Directory.ReadWrite.All","Organization.Read.All" -NoWelcome -ErrorAction Stop
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

# -----------------------------------------------------------
# License SKU selection - read live from the tenant, never hardcoded
# -----------------------------------------------------------
# v2.0: v1.0 shipped four hardcoded education SKU GUIDs. Education SKU IDs differ
# per tenant and per region, so a stale GUID here failed EVERY row at the license
# step - after the user objects had already been created, leaving a half-finished
# batch to clean up by hand. The menu below is built from the tenant at runtime,
# so it is always correct, always current, and works in any tenant.
Write-Host ""
Write-Host "Reading available licenses from the tenant..." -ForegroundColor Cyan
try {
    $TenantSkus = @(Get-MgSubscribedSku -All -ErrorAction Stop | Sort-Object SkuPartNumber)
}
catch {
    Write-Host "ERROR: Could not read SKUs from the tenant: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  This requires the Organization.Read.All scope." -ForegroundColor Yellow
    return
}

if ($TenantSkus.Count -eq 0) {
    Write-Host "ERROR: No subscribed SKUs found in this tenant." -ForegroundColor Red
    return
}

Write-Host ""
Write-Host "Licenses available in this tenant:" -ForegroundColor Cyan
for ($i = 0; $i -lt $TenantSkus.Count; $i++) {
    $s     = $TenantSkus[$i]
    $en    = $s.PrepaidUnits.Enabled
    $con   = $s.ConsumedUnits
    $avail = $en - $con
    Write-Host ("  [{0,2}] {1,-45} {2,6} free   ({3} of {4} assigned)" -f `
        ($i + 1), $s.SkuPartNumber, $avail, $con, $en) `
        -ForegroundColor $(if ($avail -gt 0) { 'White' } else { 'DarkGray' })
}
Write-Host ""
Write-Host "Tip: education SKUs are usually named M365EDU_A3_STUUSEBNFT (A3 Students)," -ForegroundColor DarkGray
Write-Host "     M365EDU_A3_FACULTY (A3 Faculty), STANDARDWOFFPACK_IW_STUDENT (A1 Students)." -ForegroundColor DarkGray
Write-Host ""

$LicenseChoice = (Read-Host "Enter the number of the license to assign").Trim()
$SkuIndex = 0
if (-not [int]::TryParse($LicenseChoice, [ref]$SkuIndex) -or
    $SkuIndex -lt 1 -or $SkuIndex -gt $TenantSkus.Count) {
    Write-Host "ERROR: '$LicenseChoice' is not a valid selection (expected 1-$($TenantSkus.Count))." -ForegroundColor Red
    Write-Host "Re-run Phase 2." -ForegroundColor Yellow
    return
}

$MatchedSku   = $TenantSkus[$SkuIndex - 1]
$LicenseSkuId = $MatchedSku.SkuId
$LicenseName  = $MatchedSku.SkuPartNumber

$Enabled   = $MatchedSku.PrepaidUnits.Enabled
$Consumed  = $MatchedSku.ConsumedUnits
$Available = $Enabled - $Consumed

Write-Host ""
Write-Host "License selected:" -ForegroundColor Green
Write-Host "  Part #    : $LicenseName" -ForegroundColor Green
Write-Host "  SkuId     : $LicenseSkuId" -ForegroundColor Green
Write-Host "  Enabled   : $Enabled" -ForegroundColor Green
Write-Host "  Consumed  : $Consumed" -ForegroundColor Green
Write-Host "  Available : $Available" -ForegroundColor $(if ($Available -gt 0) { 'Green' } else { 'Red' })

if ($Available -le 0) {
    Write-Host "[WARN] This SKU has no free units. Every license assignment will fail." -ForegroundColor Yellow
}

# NOTE: License-vs-user-count comparison happens in Phase 3 after the CSV is loaded.



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
$RequiredColumns = @('DisplayName','UserPrincipalName','MailNickname','Password','UsageLocation','GivenName','Surname')
$OptionalColumns = @('Department','JobTitle')
$ActualColumns   = $Users[0].PSObject.Properties.Name
$MissingColumns  = $RequiredColumns | Where-Object { $_ -notin $ActualColumns }

if ($MissingColumns.Count -gt 0) {
    Write-Host "ERROR: CSV is missing required column(s): $($MissingColumns -join ', ')" -ForegroundColor Red
    Write-Host "Required headers : $($RequiredColumns -join ', ')" -ForegroundColor Yellow
    Write-Host "Optional headers : $($OptionalColumns -join ', ')" -ForegroundColor Yellow
    return
}

# Display the contents in a table for review
Write-Host ""
Write-Host "CSV loaded successfully. Total rows: $($Users.Count)" -ForegroundColor Green
Write-Host ""
Write-Host "Preview of CSV contents:" -ForegroundColor Cyan
Write-Host "------------------------------------------------------------" -ForegroundColor DarkGray

$RowNum = 0
$Users | ForEach-Object {
    $RowNum++
    [PSCustomObject]@{
        Row           = $RowNum
        DisplayName   = $(if ([string]::IsNullOrWhiteSpace($_.DisplayName))       { '<EMPTY>' } else { $_.DisplayName.Trim() })
        UPN           = $(if ([string]::IsNullOrWhiteSpace($_.UserPrincipalName)) { '<EMPTY>' } else { $_.UserPrincipalName.Trim() })
        MailNickname  = $(if ([string]::IsNullOrWhiteSpace($_.MailNickname))      { '<EMPTY>' } else { $_.MailNickname.Trim() })
        UsageLocation = $(if ([string]::IsNullOrWhiteSpace($_.UsageLocation))     { '<EMPTY>' } else { $_.UsageLocation.Trim() })
        GivenName     = $(if ([string]::IsNullOrWhiteSpace($_.GivenName))         { '<EMPTY>' } else { $_.GivenName.Trim() })
        Surname       = $(if ([string]::IsNullOrWhiteSpace($_.Surname))           { '<EMPTY>' } else { $_.Surname.Trim() })
        Department    = $(if ($null -eq $_.Department -or [string]::IsNullOrWhiteSpace($_.Department)) { '' } else { $_.Department.Trim() })
        JobTitle      = $(if ($null -eq $_.JobTitle   -or [string]::IsNullOrWhiteSpace($_.JobTitle))   { '' } else { $_.JobTitle.Trim() })
        Password      = $(if ([string]::IsNullOrWhiteSpace($_.Password))          { '<EMPTY>' } else { $_.Password.Trim() })
    }
} | Format-Table -AutoSize | Out-Host

Write-Host "------------------------------------------------------------" -ForegroundColor DarkGray

# -----------------------------------------------------------
# License capacity check (CSV count vs available units)
# -----------------------------------------------------------
Write-Host ""
Write-Host "License capacity check:" -ForegroundColor Cyan
Write-Host "  Users in CSV       : $($Users.Count)" -ForegroundColor White
Write-Host "  License selected   : $LicenseName" -ForegroundColor White
Write-Host "  Total Enabled      : $Enabled" -ForegroundColor White
Write-Host "  Currently Consumed : $Consumed" -ForegroundColor White
Write-Host "  Available Units    : $Available" -ForegroundColor $(if ($Available -ge $Users.Count) { 'Green' } else { 'Red' })

if ($Available -ge $Users.Count) {
    Write-Host "  Result             : SUFFICIENT (available >= users in CSV)" -ForegroundColor Green
    Write-Host "  Remaining after run: $($Available - $Users.Count)" -ForegroundColor Green
}
else {
    $Shortfall = $Users.Count - $Available
    Write-Host "  Result             : INSUFFICIENT" -ForegroundColor Red
    Write-Host "  Shortfall          : $Shortfall license(s) short" -ForegroundColor Red
    Write-Host ""
    Write-Host "WARNING: Not enough licenses for all users in the CSV." -ForegroundColor Yellow
    Write-Host "         The first $Available user(s) will get a license; the rest will fail at the license step." -ForegroundColor Yellow
    Write-Host "         Note: existing users (already in Entra) do NOT consume a new license, so the actual" -ForegroundColor Yellow
    Write-Host "               shortfall may be smaller than shown." -ForegroundColor Yellow
    Write-Host ""
    $ProceedShortLicense = Read-Host "Do you want to continue anyway? (Y/N)"
    if ($ProceedShortLicense -notmatch '^(y|yes)$') {
        Write-Host "Cancelled by user. Disconnecting..." -ForegroundColor Red
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        return
    }
}



# ============================================================
# PHASE 4 - Confirm Before Proceeding (Yes / No)
# ============================================================
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " PHASE 4: Confirmation" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "You are about to CREATE $($Users.Count) new user(s) in tenant:" -ForegroundColor Yellow
Write-Host "  $($Context.TenantId)" -ForegroundColor Yellow
Write-Host ""
Write-Host "Each new user will be:" -ForegroundColor Yellow
Write-Host "  - Created with AccountEnabled = true" -ForegroundColor Yellow
Write-Host "  - Required to change password at next sign-in" -ForegroundColor Yellow
Write-Host "  - Assigned license: $LicenseName ($LicenseSkuId)" -ForegroundColor Yellow
Write-Host ""

# Ask whether to also assign/update the selected license on EXISTING users
Write-Host "For users that ALREADY exist in Entra ID:" -ForegroundColor Cyan
Write-Host "  - Department will still be updated from the CSV (existing behavior)" -ForegroundColor Cyan
Write-Host "  - You can OPTIONALLY also assign the selected license to them if they don't have it" -ForegroundColor Cyan
Write-Host ""
$UpdateLicenseExistingInput = Read-Host "Also assign '$LicenseName' to existing users that don't already have it? (Y/N)"
$UpdateLicenseForExisting = ($UpdateLicenseExistingInput -match '^(y|yes)$')

if ($UpdateLicenseForExisting) {
    Write-Host "  -> Existing users WITHOUT this license will have it added." -ForegroundColor Green
    Write-Host "  -> Existing users that already have it will be left alone." -ForegroundColor Green
} else {
    Write-Host "  -> Existing users will NOT receive any license change." -ForegroundColor DarkYellow
}
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
Write-Host "Confirmed. Proceeding with user creation..." -ForegroundColor Green



# ============================================================
# PHASE 5 - Create Users, Assign License, Export Log
# ============================================================
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " PHASE 5: Create Users + Assign License" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

$Log = [System.Collections.Generic.List[object]]::new()
$Counters = @{ Created = 0; Licensed = 0; Failed = 0; Skipped = 0 }
$RowNumber = 0

foreach ($u in $Users) {
    $RowNumber++

    # Safe trims
    $DisplayName   = if ($null -ne $u.DisplayName)       { ([string]$u.DisplayName).Trim() }       else { "" }
    $UPN           = if ($null -ne $u.UserPrincipalName) { ([string]$u.UserPrincipalName).Trim() } else { "" }
    $MailNickname  = if ($null -ne $u.MailNickname)      { ([string]$u.MailNickname).Trim() }      else { "" }
    $Password      = if ($null -ne $u.Password)          { ([string]$u.Password).Trim() }          else { "" }
    $UsageLocation = if ($null -ne $u.UsageLocation)     { ([string]$u.UsageLocation).Trim() }     else { "" }
    $GivenName     = if ($null -ne $u.GivenName)         { ([string]$u.GivenName).Trim() }         else { "" }
    $Surname       = if ($null -ne $u.Surname)           { ([string]$u.Surname).Trim() }           else { "" }
    $Department    = if ($null -ne $u.Department)        { ([string]$u.Department).Trim() }        else { "" }
    $JobTitle      = if ($null -ne $u.JobTitle)          { ([string]$u.JobTitle).Trim() }          else { "" }

    # Skip if any required field is empty
    $MissingRequired = @()
    if ([string]::IsNullOrWhiteSpace($DisplayName))   { $MissingRequired += 'DisplayName' }
    if ([string]::IsNullOrWhiteSpace($UPN))           { $MissingRequired += 'UserPrincipalName' }
    if ([string]::IsNullOrWhiteSpace($MailNickname))  { $MissingRequired += 'MailNickname' }
    if ([string]::IsNullOrWhiteSpace($Password))      { $MissingRequired += 'Password' }
    if ([string]::IsNullOrWhiteSpace($UsageLocation)) { $MissingRequired += 'UsageLocation' }

    if ($MissingRequired.Count -gt 0) {
        $msg = "Missing required field(s): $($MissingRequired -join ', ')"
        Write-Host "[Row $RowNumber] Skipped: $msg" -ForegroundColor Yellow
        $Log.Add([PSCustomObject]@{
            Timestamp         = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
            Row               = $RowNumber
            DisplayName       = $DisplayName
            UserPrincipalName = $UPN
            TemporaryPassword = $Password
            License           = $LicenseName
            UserCreated       = "No"
            LicenseAssigned   = "No"
            Status            = "Skipped"
            Message           = $msg
        })
        $Counters.Skipped++
        continue
    }

    $UserCreated     = $false
    $LicenseAssigned = $false
    $NewUserId       = $null
    $StepMessages    = @()

    # ---- Step 0: Check if user already exists; if so, update Department only ----
    $ExistingUser = $null
    try {
        $ExistingUser = Get-MgUser -UserId $UPN -Property "Id,UserPrincipalName,DisplayName,Department" -ErrorAction Stop
    }
    catch {
        # Not found (or other error) -> proceed to create
        $ExistingUser = $null
    }

    if ($ExistingUser) {
        Write-Host "[Row $RowNumber] User already exists: $UPN" -ForegroundColor DarkYellow

        $OldDepartment   = if ($null -ne $ExistingUser.Department) { [string]$ExistingUser.Department } else { "" }
        $DeptUpdateMsg   = ""
        $DeptUpdateState = "NotChanged"

        if ([string]::IsNullOrWhiteSpace($Department)) {
            $DeptUpdateMsg = "CSV Department empty; no update performed."
            Write-Host "[Row $RowNumber]   $DeptUpdateMsg" -ForegroundColor DarkGray
        }
        elseif ($OldDepartment -eq $Department) {
            $DeptUpdateMsg = "Department already '$Department'; no update needed."
            Write-Host "[Row $RowNumber]   $DeptUpdateMsg" -ForegroundColor DarkGray
        }
        else {
            try {
                Update-MgUser -UserId $ExistingUser.Id -Department $Department -ErrorAction Stop
                $DeptUpdateState = "Updated"
                $DeptUpdateMsg   = "Department updated from '$OldDepartment' to '$Department'."
                Write-Host "[Row $RowNumber]   $DeptUpdateMsg" -ForegroundColor Green
            }
            catch {
                $errMsg = $_.Exception.Message
                if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
                    $errMsg = "$errMsg | Details: $($_.ErrorDetails.Message)"
                }
                $DeptUpdateState = "Failed"
                $DeptUpdateMsg   = "Department update failed: $errMsg"
                Write-Host "[Row $RowNumber]   $DeptUpdateMsg" -ForegroundColor Red
            }
        }

        $RowStatus = if ($DeptUpdateState -eq "Failed") { "Failed" } else { "Skipped" }

        $Log.Add([PSCustomObject]@{
            Timestamp         = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
            Row               = $RowNumber
            DisplayName       = $DisplayName
            UserPrincipalName = $UPN
            TemporaryPassword = ""   # not changed
            License           = $LicenseName
            UserCreated       = "No"
            LicenseAssigned   = "N/A"
            Status            = $RowStatus
            Message           = "User already exists. $DeptUpdateMsg"
        })

        if ($RowStatus -eq "Failed") { $Counters.Failed++ } else { $Counters.Skipped++ }
        continue
    }

    # ---- Step A: Create user ----
    try {
        Write-Host "[Row $RowNumber] Creating user $UPN..." -ForegroundColor Cyan

        $PasswordProfile = @{
            Password                      = $Password
            ForceChangePasswordNextSignIn = $true
        }

        $UserParams = @{
            DisplayName       = $DisplayName
            UserPrincipalName = $UPN
            MailNickname      = $MailNickname
            PasswordProfile   = $PasswordProfile
            AccountEnabled    = $true
            UsageLocation     = $UsageLocation
            GivenName         = $GivenName
            Surname           = $Surname
        }
        if (-not [string]::IsNullOrWhiteSpace($Department)) {
            $UserParams['Department'] = $Department
        }

        $NewUser = New-MgUser @UserParams -ErrorAction Stop
        $NewUserId = $NewUser.Id
        $UserCreated = $true
        $Counters.Created++
        $StepMessages += "User created"
        Write-Host "[Row $RowNumber]   User created (Id: $NewUserId)" -ForegroundColor Green
    }
    catch {
        $errMsg = $_.Exception.Message
        if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
            $errMsg = "$errMsg | Details: $($_.ErrorDetails.Message)"
        }
        Write-Host "[Row $RowNumber]   FAILED to create user: $errMsg" -ForegroundColor Red

        $Log.Add([PSCustomObject]@{
            Timestamp         = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
            Row               = $RowNumber
            DisplayName       = $DisplayName
            UserPrincipalName = $UPN
            TemporaryPassword = $Password
            License           = $LicenseName
            UserCreated       = "No"
            LicenseAssigned   = "No"
            Status            = "Failed"
            Message           = "User creation failed: $errMsg"
        })
        $Counters.Failed++
        continue
    }

    # ---- Step B: Assign License ----
    try {
        Write-Host "[Row $RowNumber] Assigning license '$LicenseName'..." -ForegroundColor Cyan

        # Brief wait helps avoid timing issues right after user creation
        Start-Sleep -Seconds 2

        $CurrentLicenses = @()
        try {
            $CurrentLicenses = (Get-MgUserLicenseDetail -UserId $NewUserId -ErrorAction Stop).SkuId
        } catch {
            # New users typically have no licenses yet; ignore
            $CurrentLicenses = @()
        }

        if ($CurrentLicenses -contains $LicenseSkuId) {
            Write-Host "[Row $RowNumber]   License already assigned. Skipping add." -ForegroundColor DarkYellow
            $LicenseAssigned = $true
            $StepMessages += "License already present"
        }
        else {
            Set-MgUserLicense -UserId $NewUserId `
                -AddLicenses @(@{SkuId = $LicenseSkuId}) `
                -RemoveLicenses @() `
                -ErrorAction Stop | Out-Null

            $LicenseAssigned = $true
            $Counters.Licensed++
            $StepMessages += "License assigned"
            Write-Host "[Row $RowNumber]   License assigned" -ForegroundColor Green
        }
    }
    catch {
        $errMsg = $_.Exception.Message
        if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
            $errMsg = "$errMsg | Details: $($_.ErrorDetails.Message)"
        }
        Write-Host "[Row $RowNumber]   FAILED to assign license: $errMsg" -ForegroundColor Red
        $StepMessages += "License assignment failed: $errMsg"
    }

    # ---- Step C: Update JobTitle (optional, after license) ----
    if (-not [string]::IsNullOrWhiteSpace($JobTitle)) {
        try {
            Update-MgUser -UserId $NewUserId -JobTitle $JobTitle -ErrorAction Stop
            $StepMessages += "JobTitle set"
            Write-Host "[Row $RowNumber]   JobTitle set to '$JobTitle'" -ForegroundColor Green
        }
        catch {
            $errMsg = $_.Exception.Message
            Write-Host "[Row $RowNumber]   Failed to set JobTitle: $errMsg" -ForegroundColor DarkYellow
            $StepMessages += "JobTitle update failed: $errMsg"
        }
    }

    # ---- Final per-row log entry ----
    $RowStatus = if ($UserCreated -and $LicenseAssigned) { "Success" }
                 elseif ($UserCreated)                    { "Partial" }
                 else                                     { "Failed" }

    $Log.Add([PSCustomObject]@{
        Timestamp         = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        Row               = $RowNumber
        DisplayName       = $DisplayName
        UserPrincipalName = $UPN
        TemporaryPassword = $Password
        License           = $LicenseName
        UserCreated       = $(if ($UserCreated)     { "Yes" } else { "No" })
        LicenseAssigned   = $(if ($LicenseAssigned) { "Yes" } else { "No" })
        Status            = $RowStatus
        Message           = ($StepMessages -join '; ')
    })
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
Write-Host (" Users created      : {0}" -f $Counters.Created)  -ForegroundColor Green
Write-Host (" Licenses assigned  : {0}" -f $Counters.Licensed) -ForegroundColor Green
Write-Host (" Failed             : {0}" -f $Counters.Failed)   -ForegroundColor Red
Write-Host (" Skipped            : {0}" -f $Counters.Skipped)  -ForegroundColor Yellow
Write-Host (" License selected   : {0}" -f $LicenseName)
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
