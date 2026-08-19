<#
============================================================
Bulk Extract M365 Storage Report - Phased Execution
============================================================
Description:
- Designed for copy-paste execution into a PowerShell terminal
- Runs in 5 distinct phases with user confirmation
- Pulls tenant-wide storage from Microsoft Graph (period=D30)
- Consolidates into a 3-row / 5-column CSV (one row per service)

Endpoints pulled (in parallel, PS7+):
  1. getOneDriveUsageAccountDetail(period='D30')   -> per-user OneDrive
  2. getSharePointSiteUsageDetail(period='D30')    -> per-site SharePoint
  3. getMailboxUsageDetail(period='D30')           -> per-mailbox Exchange

Output columns (exact 7 columns, single row):
  ONEDRIVE, EXCHANGE, SHARE POINT, CURRENT STORAGE ( <MONTH>),
  USAGE, Used Storage(GB), Total Storage(GB)

Column semantics:
  ONEDRIVE                = OneDrive used (human-readable size)
  EXCHANGE                = Exchange used (human-readable size)
  SHARE POINT             = SharePoint used (human-readable size)
  CURRENT STORAGE ( <MONTH>) = Combined used across all 3 (human-readable size)
                            <MONTH> derives from the run date; overridable at the prompt.
  USAGE                   = "<X.XX> TB of <Y.YY> TB used" (pooled TB fraction)
  Used Storage(GB)        = Combined used, numeric GB (2 decimals)
  Total Storage(GB)       = Combined quota, numeric GB (2 decimals)

Sizing notes:
- Used  = sum of "Storage Used (Byte)" across active (non-deleted) rows
- Total = sum of "Storage Allocated (Byte)" for OneDrive / SharePoint
       = sum of "Prohibit Send/Receive Quota (Byte)" for Exchange
- Binary GB (1 GB = 1024^3 bytes) - matches Microsoft admin center portal
- Deleted accounts/sites/mailboxes are excluded (Is Deleted = False)

Requirements:
- PowerShell 7+
- Microsoft.Graph.Authentication module
- Admin role: Reports Reader / Global Reader / Global Admin
- Scopes: Reports.Read.All, Organization.Read.All

How to use:
- Copy-paste each phase into a PowerShell 7+ terminal
- OR paste the entire script at once

.AUTHOR         Generated with Claude for Rakso CT Education IT.
.VERSION        2.0
.DATE           2026-08-19
.REQUIREMENTS   PowerShell 7+ (parallel pull); Microsoft.Graph.Authentication.
.PERMISSIONS    Reports.Read.All, Organization.Read.All (both read-only).
.SAFETY         Read-only. Pulls usage reports only; nothing is written to the tenant.
.CHANGELOG      v1.0 - Initial release.
                v2.0 - Removed the hardcoded tenant GUID and the hardcoded output
                       folder. Tenant ID is prompted and GUID-validated at runtime;
                       the output folder defaults to the script's own folder and stays
                       overridable. The snapshot column label, previously frozen at
                       "CURRENT STORAGE ( JUNE)", now derives from the run month and
                       can be overridden - the old build mislabelled the column on
                       every run outside June.
============================================================
#>


# ============================================================
# PHASE 1 - Configure Paths + Period
# ============================================================
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " PHASE 1: Configure Paths + Period" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

# Helper: strip invisible/Unicode formatting chars from pasted paths
function Remove-InvisibleChars {
    param([string]$s)
    if ([string]::IsNullOrEmpty($s)) { return $s }
    return ($s -replace '[\u200B-\u200F\u202A-\u202E\u2066-\u2069\uFEFF\u00A0\u202F\u1680\u2000-\u200A\u205F\u3000]', '')
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

$DefaultOutputFolder = Join-Path $ScriptDir 'Reports'
$Period = "D30"

Write-Host ""
Write-Host "Default output folder:" -ForegroundColor Gray
Write-Host "  $DefaultOutputFolder" -ForegroundColor DarkGray
$InputFolder = Read-Host "Enter output folder (or press Enter to use default)"
if ([string]::IsNullOrWhiteSpace($InputFolder)) {
    $OutputFolder = $DefaultOutputFolder
} else {
    $OutputFolder = (Remove-InvisibleChars $InputFolder).Trim('"').Trim("'").Trim()
}

# Create folder if missing
if (-not (Test-Path -LiteralPath $OutputFolder)) {
    try {
        New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null
        Write-Host "Created output folder: $OutputFolder" -ForegroundColor DarkGray
    } catch {
        Write-Host "ERROR: Could not create folder '$OutputFolder': $($_.Exception.Message)" -ForegroundColor Red
        return
    }
}

# Timestamped output path
$TimeStamp  = Get-Date -Format "yyyyMMdd_HHmmss"
$OutputPath = Join-Path -Path $OutputFolder -ChildPath "M365_StorageReport_${Period}_$TimeStamp.csv"

# Snapshot column label. v1.0 froze this at "CURRENT STORAGE ( JUNE)", so every
# report run outside June shipped with a wrong month on the column. It now derives
# from the run date and can still be overridden (e.g. to label a back-dated pull).
$DefaultMonthLabel = (Get-Date).ToString('MMMM').ToUpperInvariant()
Write-Host ""
$InputMonthLabel = Read-Host "Snapshot month label for the storage column (Enter = $DefaultMonthLabel)"
$MonthLabel = if ([string]::IsNullOrWhiteSpace($InputMonthLabel)) { $DefaultMonthLabel }
              else { (Remove-InvisibleChars $InputMonthLabel).Trim().ToUpperInvariant() }
$CurrentStorageCol = "CURRENT STORAGE ( $MonthLabel)"

Write-Host ""
Write-Host "Configuration:" -ForegroundColor Green
Write-Host "  Period          : $Period (past 30 days)" -ForegroundColor Green
Write-Host "  Snapshot column : $CurrentStorageCol" -ForegroundColor Green
Write-Host "  Output CSV      : $OutputPath" -ForegroundColor Green


# ============================================================
# PHASE 2 - Connect to Microsoft Graph (Tenant ID + Sign-In Mode)
# ============================================================
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " PHASE 2: Connect to Microsoft Graph" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

# Require PS7+
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host ""
    Write-Host "ERROR: PowerShell 7+ is required (parallel processing)." -ForegroundColor Red
    Write-Host "Current version: $($PSVersionTable.PSVersion)" -ForegroundColor Red
    return
}

# Ensure Microsoft.Graph.Authentication is present
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
    Write-Host ""
    Write-Host "ERROR: Microsoft.Graph.Authentication module not installed." -ForegroundColor Red
    Write-Host "Install with:" -ForegroundColor Yellow
    Write-Host "  Install-Module Microsoft.Graph -Scope CurrentUser -Force" -ForegroundColor Yellow
    return
}
try {
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
} catch {
    Write-Host "ERROR: Failed to import module: $($_.Exception.Message)" -ForegroundColor Red
    return
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

# Sign-in mode
Write-Host ""
Write-Host "Sign-in modes:" -ForegroundColor Yellow
Write-Host "  [1] Browser / DeviceCode flow (default)" -ForegroundColor DarkGray
Write-Host "  [2] Interactive popup" -ForegroundColor DarkGray
Write-Host "  [3] Reuse existing session" -ForegroundColor DarkGray
$SignInMode = Read-Host "Select sign-in mode (1/2/3, Enter=1)"
if ([string]::IsNullOrWhiteSpace($SignInMode)) { $SignInMode = "1" }

$Scopes = @("Reports.Read.All","Organization.Read.All")

try {
    switch ($SignInMode) {
        "2" {
            Write-Host ""
            Write-Host "Connecting via Interactive popup..." -ForegroundColor Cyan
            Connect-MgGraph -TenantId $TenantId -Scopes $Scopes -NoWelcome -ErrorAction Stop
        }
        "3" {
            $ExistingCtx = Get-MgContext
            if (-not $ExistingCtx) {
                Write-Host ""
                Write-Host "No existing session. Falling back to DeviceCode." -ForegroundColor Yellow
                Connect-MgGraph -TenantId $TenantId -Scopes $Scopes -UseDeviceCode -NoWelcome -ErrorAction Stop
            } else {
                Write-Host ""
                Write-Host "Reusing existing session: $($ExistingCtx.Account)" -ForegroundColor Green
            }
        }
        default {
            Write-Host ""
            Write-Host "Connecting via DeviceCode flow..." -ForegroundColor Cyan
            Connect-MgGraph -TenantId $TenantId -Scopes $Scopes -UseDeviceCode -NoWelcome -ErrorAction Stop
        }
    }

    $Context = Get-MgContext
    if (-not $Context) { throw "No Graph context after Connect-MgGraph." }

    Write-Host ""
    Write-Host "Connected successfully." -ForegroundColor Green
    Write-Host "  Account : $($Context.Account)" -ForegroundColor DarkGray
    Write-Host "  Tenant  : $($Context.TenantId)" -ForegroundColor DarkGray
    Write-Host "  Scopes  : $($Context.Scopes -join ', ')" -ForegroundColor DarkGray
}
catch {
    Write-Host "ERROR: Failed to connect: $($_.Exception.Message)" -ForegroundColor Red
    return
}


# ============================================================
# PHASE 3 - Preview Endpoints + Scope Check
# ============================================================
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " PHASE 3: Preview Endpoints + Scope Check" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

$Endpoints = @(
    [PSCustomObject]@{ Name="OneDrive";   Uri="https://graph.microsoft.com/v1.0/reports/getOneDriveUsageAccountDetail(period='$Period')" }
    [PSCustomObject]@{ Name="SharePoint"; Uri="https://graph.microsoft.com/v1.0/reports/getSharePointSiteUsageDetail(period='$Period')" }
    [PSCustomObject]@{ Name="Exchange";   Uri="https://graph.microsoft.com/v1.0/reports/getMailboxUsageDetail(period='$Period')" }
)

Write-Host ""
Write-Host "Reports to pull (parallel, ThrottleLimit=3):" -ForegroundColor Yellow
$Endpoints | Format-Table Name, Uri -AutoSize | Out-Host

# Scope validation
$RequiredScopes = @("Reports.Read.All","Organization.Read.All")
$GrantedScopes  = @($Context.Scopes)
$MissingScopes  = $RequiredScopes | Where-Object { $_ -notin $GrantedScopes }

if ($MissingScopes.Count -gt 0) {
    Write-Host "WARNING: Missing scopes: $($MissingScopes -join ', ')" -ForegroundColor Yellow
    Write-Host "Some reports may fail. Reconnect with the missing scopes if needed." -ForegroundColor Yellow
} else {
    Write-Host "All required scopes granted." -ForegroundColor Green
}


# ============================================================
# PHASE 4 - Confirmation
# ============================================================
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " PHASE 4: Confirmation" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

Write-Host ""
Write-Host "Summary:" -ForegroundColor Yellow
Write-Host "  Tenant ID    : $TenantId" -ForegroundColor DarkGray
Write-Host "  Account      : $($Context.Account)" -ForegroundColor DarkGray
Write-Host "  Period       : $Period" -ForegroundColor DarkGray
Write-Host "  Endpoints    : $($Endpoints.Count)" -ForegroundColor DarkGray
Write-Host "  Output CSV   : $OutputPath" -ForegroundColor DarkGray
Write-Host ""

$Confirm = Read-Host "Proceed to pull reports? (Y/N)"
if ($Confirm -notmatch '^(y|yes)$') {
    Write-Host ""
    Write-Host "Operation cancelled by user. No changes were made." -ForegroundColor Red
    Write-Host "Disconnecting..." -ForegroundColor DarkGray
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    return
}


# ============================================================
# PHASE 5 - Parallel Pull + Aggregate + Export
# ============================================================
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " PHASE 5: Pull Reports in Parallel + Export" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

# ------------------------------------------------------------
# Acquire raw access token for parallel REST calls.
# ------------------------------------------------------------
try {
    $TokenResp = Invoke-MgGraphRequest `
        -Method GET `
        -Uri "https://graph.microsoft.com/v1.0/organization" `
        -OutputType HttpResponseMessage `
        -ErrorAction Stop
    $AccessToken = $TokenResp.RequestMessage.Headers.Authorization.Parameter
    if ([string]::IsNullOrWhiteSpace($AccessToken)) { throw "Empty access token" }
    Write-Host "Access token acquired for parallel calls." -ForegroundColor Green
} catch {
    Write-Host "ERROR: Could not acquire access token: $($_.Exception.Message)" -ForegroundColor Red
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    return
}

Write-Host ""
Write-Host "Pulling $($Endpoints.Count) endpoints in parallel..." -ForegroundColor Cyan

$Results = $Endpoints | ForEach-Object -Parallel {
    $ep         = $_
    $token      = $using:AccessToken
    $maxRetries = 5
    $headers    = @{ Authorization = "Bearer $token" }

    $success = $false
    $errMsg  = $null
    $content = $null

    for ($i = 0; $i -lt $maxRetries; $i++) {
        try {
            $resp = Invoke-WebRequest -Uri $ep.Uri -Headers $headers -UseBasicParsing -ErrorAction Stop

            # Graph returns CSV as byte[] (Content-Type: application/octet-stream).
            # Normalize to UTF-8 string + strip BOM.
            if ($resp.Content -is [byte[]]) {
                $content = [System.Text.Encoding]::UTF8.GetString($resp.Content)
            } else {
                $content = [string]$resp.Content
            }
            if ($content.Length -gt 0 -and $content[0] -eq [char]0xFEFF) {
                $content = $content.Substring(1)
            }

            $success = $true
            break
        } catch {
            $statusCode = $null
            try { $statusCode = [int]$_.Exception.Response.StatusCode } catch {}

            if ($statusCode -eq 429 -or $statusCode -eq 503) {
                $retryAfter = 0
                try { $retryAfter = [int]$_.Exception.Response.Headers.RetryAfter.Delta.TotalSeconds } catch {}
                if ($retryAfter -le 0) { $retryAfter = [int][math]::Min([math]::Pow(2, $i + 1), 60) }
                Start-Sleep -Seconds $retryAfter
                continue
            }

            $errMsg = "$statusCode :: $($_.Exception.Message)"
            if ($i -lt $maxRetries - 1) {
                Start-Sleep -Seconds ([int][math]::Pow(2, $i))
            } else {
                break
            }
        }
    }

    [PSCustomObject]@{
        Name    = $ep.Name
        Success = $success
        Content = $content
        Error   = $errMsg
    }
} -ThrottleLimit 3

# Report per-endpoint status
Write-Host ""
foreach ($r in $Results) {
    if ($r.Success) {
        Write-Host ("  [OK]   {0} ({1} bytes)" -f $r.Name, $r.Content.Length) -ForegroundColor Green
    } else {
        Write-Host ("  [FAIL] {0} :: {1}" -f $r.Name, $r.Error) -ForegroundColor Red
    }
}

# ------------------------------------------------------------
# Parse each response into rows
# ------------------------------------------------------------
$OneDriveData   = $null
$SharePointData = $null
$ExchangeData   = $null

foreach ($r in $Results) {
    if (-not $r.Success) { continue }
    try {
        switch ($r.Name) {
            "OneDrive"   { $OneDriveData   = $r.Content | ConvertFrom-Csv }
            "SharePoint" { $SharePointData = $r.Content | ConvertFrom-Csv }
            "Exchange"   { $ExchangeData   = $r.Content | ConvertFrom-Csv }
        }
    } catch {
        Write-Host ("Parse error for {0}: {1}" -f $r.Name, $_.Exception.Message) -ForegroundColor Red
    }
}

# ------------------------------------------------------------
# Helpers: safe [long] sum + byte formatter
# ------------------------------------------------------------
function Get-SumBytes {
    param($Data, [string]$Property, [scriptblock]$Filter = { $true })
    $total = [long]0
    if (-not $Data) { return $total }
    foreach ($row in $Data) {
        if (-not (& $Filter $row)) { continue }
        $v = $row.$Property
        if ([string]::IsNullOrWhiteSpace($v)) { continue }
        $n = [long]0
        if ([long]::TryParse($v, [ref]$n)) { $total += $n }
    }
    return $total
}

function Format-Size {
    param([long]$Bytes)
    if ($Bytes -ge 1TB) { return ("{0:N2} TB" -f ($Bytes / 1TB)) }
    if ($Bytes -ge 1GB) { return ("{0:N2} GB" -f ($Bytes / 1GB)) }
    if ($Bytes -ge 1MB) { return ("{0:N2} MB" -f ($Bytes / 1MB)) }
    if ($Bytes -ge 1KB) { return ("{0:N2} KB" -f ($Bytes / 1KB)) }
    return "$Bytes B"
}

function ConvertTo-GB {
    param([long]$Bytes)
    return [math]::Round($Bytes / 1GB, 2)
}

function Get-UsagePct {
    param([long]$Used, [long]$Total)
    if ($Total -le 0) { return "0.00%" }
    return ("{0:N2}%" -f (($Used / $Total) * 100))
}

function Format-TBUsage {
    # Formats "X.XX TB of Y.YY TB used" - forces TB unit regardless of size
    param([long]$Used, [long]$Total)
    $usedTB  = [math]::Round($Used  / 1TB, 2)
    $totalTB = [math]::Round($Total / 1TB, 2)
    return ("{0:N2} TB of {1:N2} TB used" -f $usedTB, $totalTB)
}

# Filter: exclude deleted accounts/sites/mailboxes
$ActiveFilter = { param($row) ($row.'Is Deleted' -ne 'True') -and ($row.'Is Deleted' -ne 'true') }

# ------------------------------------------------------------
# Aggregate per service
# ------------------------------------------------------------

# OneDrive
$odUsed  = Get-SumBytes -Data $OneDriveData -Property 'Storage Used (Byte)'      -Filter $ActiveFilter
$odTotal = Get-SumBytes -Data $OneDriveData -Property 'Storage Allocated (Byte)' -Filter $ActiveFilter

# SharePoint
$spUsed  = Get-SumBytes -Data $SharePointData -Property 'Storage Used (Byte)'      -Filter $ActiveFilter
$spTotal = Get-SumBytes -Data $SharePointData -Property 'Storage Allocated (Byte)' -Filter $ActiveFilter

# Exchange (Total = Prohibit Send/Receive Quota)
$exUsed  = Get-SumBytes -Data $ExchangeData -Property 'Storage Used (Byte)'                    -Filter $ActiveFilter
$exTotal = Get-SumBytes -Data $ExchangeData -Property 'Prohibit Send/Receive Quota (Byte)'     -Filter $ActiveFilter

# ------------------------------------------------------------
# Build single-row 7-column output (exact header order)
# ------------------------------------------------------------
$TotalUsed  = $odUsed  + $exUsed  + $spUsed
$TotalQuota = $odTotal + $exTotal + $spTotal

$Row = [ordered]@{
    'ONEDRIVE'           = (Format-Size -Bytes $odUsed)
    'EXCHANGE'           = (Format-Size -Bytes $exUsed)
    'SHARE POINT'        = (Format-Size -Bytes $spUsed)
    $CurrentStorageCol   = (Format-Size -Bytes $TotalUsed)
    'USAGE'              = (Format-TBUsage -Used $TotalUsed -Total $TotalQuota)
    'Used Storage(GB)'   = (ConvertTo-GB -Bytes $TotalUsed)
    'Total Storage(GB)'  = (ConvertTo-GB -Bytes $TotalQuota)
}
$Output = [PSCustomObject]$Row

# Preview (list format so long values wrap cleanly)
Write-Host ""
Write-Host "Consolidated result:" -ForegroundColor Yellow
$Output | Format-List | Out-Host

# Diagnostic counts
Write-Host "Row counts (active only):" -ForegroundColor DarkGray
Write-Host ("  OneDrive   : {0}" -f (@($OneDriveData   | Where-Object { $_.'Is Deleted' -ne 'True' }).Count)) -ForegroundColor DarkGray
Write-Host ("  SharePoint : {0}" -f (@($SharePointData | Where-Object { $_.'Is Deleted' -ne 'True' }).Count)) -ForegroundColor DarkGray
Write-Host ("  Exchange   : {0}" -f (@($ExchangeData   | Where-Object { $_.'Is Deleted' -ne 'True' }).Count)) -ForegroundColor DarkGray

# Export CSV
try {
    $Output | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host ""
    Write-Host "CSV exported successfully:" -ForegroundColor Green
    Write-Host "  $OutputPath" -ForegroundColor Green
} catch {
    Write-Host "ERROR: Export failed: $($_.Exception.Message)" -ForegroundColor Red
}

# Summary
$OkCount = ($Results | Where-Object Success).Count
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " Summary" -ForegroundColor Cyan
Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
Write-Host (" Period               : {0}" -f $Period)
Write-Host (" Endpoints succeeded  : {0}/{1}" -f $OkCount, $Endpoints.Count)
Write-Host (" Output file          : {0}" -f $OutputPath)
Write-Host (" End Time             : {0}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
Write-Host "============================================================" -ForegroundColor Cyan

# Disconnect
try {
    Disconnect-MgGraph -ErrorAction Stop | Out-Null
    Write-Host "Disconnected from Microsoft Graph." -ForegroundColor DarkGray
} catch {
    Write-Host "Warning: Disconnect-MgGraph failed: $($_.Exception.Message)" -ForegroundColor Yellow
}
