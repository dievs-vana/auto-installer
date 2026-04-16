#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Microsoft 365 Silent Installer with System Preparation

.USAGE
    Run from PowerShell (Admin):
    irm https://raw.githubusercontent.com/dievs-vana/auto-installer/main/Run-Setup.ps1 | iex
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ---------------------------------------------
#  BANNER
# ---------------------------------------------
Clear-Host
Write-Host ""
Write-Host "  +======================================================+" -ForegroundColor Cyan
Write-Host "  |         Microsoft 365 Auto-Setup & Provisioner       |" -ForegroundColor Cyan
Write-Host "  |              System Prep + Silent Install             |" -ForegroundColor Cyan
Write-Host "  +======================================================+" -ForegroundColor Cyan
Write-Host ""

# ---------------------------------------------
#  STATIC CONFIG  (no credentials here)
# ---------------------------------------------
$CONFIG = @{
    RepoBase         = "https://raw.githubusercontent.com/dievs-vana/auto-installer/main"
    WorkDir          = "$env:TEMP\M365Setup"
    LogFile          = "$env:TEMP\M365Setup\setup.log"
    SMTPServer       = "smtp.gmail.com"
    SMTPPort         = 587
    SkipDebloat      = $false
    SkipRestorePoint = $false
    SkipHWReport     = $false
    # Credentials filled in at runtime (see Step 0)
    ReportEmail      = ""
    SMTPUser         = ""
    SMTPPass         = ""
}

# ---------------------------------------------
#  HELPERS
# ---------------------------------------------
function Log {
    param([string]$Msg, [string]$Level = "INFO")
    $ts   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$ts][$Level] $Msg"
    if (Test-Path (Split-Path $CONFIG.LogFile)) {
        Add-Content -Path $CONFIG.LogFile -Value $line -Encoding UTF8
    }
    switch ($Level) {
        "INFO"    { Write-Host "  [..] $Msg" -ForegroundColor White }
        "OK"      { Write-Host "  [[OK]] $Msg" -ForegroundColor Green }
        "WARN"    { Write-Host "  [!] $Msg" -ForegroundColor Yellow }
        "ERROR"   { Write-Host "  [[FAIL]] $Msg" -ForegroundColor Red }
        "SECTION" {
            Write-Host ""
            Write-Host "  -- $Msg --" -ForegroundColor Cyan
        }
    }
}

function Ensure-WorkDir {
    if (-not (Test-Path $CONFIG.WorkDir)) {
        New-Item -ItemType Directory -Path $CONFIG.WorkDir -Force | Out-Null
    }
    New-Item -ItemType Directory -Force -Path "$($CONFIG.WorkDir)\Modules" | Out-Null
    New-Item -ItemType Directory -Force -Path "$($CONFIG.WorkDir)\Config"  | Out-Null
}

function Download-File {
    param([string]$RelPath)
    # Always use forward slashes in the URL
    $url  = "$($CONFIG.RepoBase)/$($RelPath -replace '\\', '/')"
    $dest = Join-Path $CONFIG.WorkDir ($RelPath -replace '/', '\')
    $dir  = Split-Path $dest
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
    Log "Downloading $RelPath..."
    try {
        Invoke-WebRequest -Uri $url -OutFile $dest -UseBasicParsing
        # Windows PowerShell can misread UTF-8 scripts without BOM; normalize .ps1 files.
        if ([System.IO.Path]::GetExtension($dest).ToLowerInvariant() -eq ".ps1") {
            $raw = [System.IO.File]::ReadAllText($dest)
            $utf8Bom = New-Object System.Text.UTF8Encoding($true)
            [System.IO.File]::WriteAllText($dest, $raw, $utf8Bom)
        }
        Log "Downloaded: $RelPath" "OK"
    } catch {
        Log "Failed to download ${RelPath}: $_" "WARN"
    }
    return $dest
}

# ---------------------------------------------
#  STEP 0 - COLLECT ALL CREDENTIALS (once)
# ---------------------------------------------
Log "Credential Collection" "SECTION"
Write-Host ""
Write-Host "  +-----------------------------------------------------+" -ForegroundColor DarkCyan
Write-Host "  |  Step 1 of 2 - Microsoft 365 Account               |" -ForegroundColor DarkCyan
Write-Host "  +-----------------------------------------------------+" -ForegroundColor DarkCyan
Write-Host ""
Write-Host "  Enter the M365 account to pre-configure after install." -ForegroundColor White
Write-Host "  Credentials stay in memory only - never written to disk." -ForegroundColor Gray
Write-Host ""

$M365User = Read-Host "  Microsoft 365 Email (UPN)"
if ($M365User -notmatch "^[^@\s]+@[^@\s]+\.[^@\s]+$") {
    Write-Host "  [[FAIL]] Invalid email format." -ForegroundColor Red
    exit 1
}
$M365Pass = Read-Host "  M365 Password" -AsSecureString
Log "M365 credentials collected." "OK"

Write-Host ""
Write-Host "  +-----------------------------------------------------+" -ForegroundColor DarkCyan
Write-Host "  |  Step 2 of 2 - Hardware Report Email (SMTP)        |" -ForegroundColor DarkCyan
Write-Host "  +-----------------------------------------------------+" -ForegroundColor DarkCyan
Write-Host ""
Write-Host "  A hardware report will be emailed after diagnostics." -ForegroundColor White
Write-Host "  Use a Gmail account with an App Password (not your real password)." -ForegroundColor Gray
Write-Host "  Get one at: https://myaccount.google.com/apppasswords" -ForegroundColor Gray
Write-Host ""

$smtpFrom = Read-Host "  Sender Gmail address"
$smtpPassRaw = Read-Host "  Gmail App Password (16 chars)" -AsSecureString
$smtpTo   = Read-Host "  Send report TO (your email)"

# Convert SMTP password SecureString -> plain for Send-MailMessage
$smtpBSTR     = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($smtpPassRaw)
$smtpPlain    = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($smtpBSTR)
[System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($smtpBSTR)

$CONFIG.SMTPUser    = $smtpFrom
$CONFIG.SMTPPass    = $smtpPlain
$CONFIG.ReportEmail = $smtpTo

Log "SMTP credentials collected." "OK"

# ---------------------------------------------
#  STEP 1 - WORKSPACE
# ---------------------------------------------
Log "Initializing workspace" "SECTION"
Ensure-WorkDir
Log "Work directory: $($CONFIG.WorkDir)" "OK"

# ---------------------------------------------
#  STEP 2 - DOWNLOAD ALL MODULES
# ---------------------------------------------
Log "Downloading setup modules" "SECTION"
$files = @(
    "Modules/1-Debloat.ps1",
    "Modules/2-RestorePoint.ps1",
    "Modules/3-HardwareReport.ps1",
    "Modules/4-InstallM365.ps1",
    "Modules/5-ConfigureM365.ps1",
    "Config/ODT-Config.xml"
)
foreach ($f in $files) { Download-File $f }

# ---------------------------------------------
#  PHASE 1 - DEBLOAT
# ---------------------------------------------
if (-not $CONFIG.SkipDebloat) {
    Log "Phase 1 - System Debloat" "SECTION"
    $s = "$($CONFIG.WorkDir)\Modules\1-Debloat.ps1"
    if (Test-Path $s) { & $s } else { Log "Debloat script not found, skipping." "WARN" }
}

# ---------------------------------------------
#  PHASE 2 - RESTORE POINT
# ---------------------------------------------
if (-not $CONFIG.SkipRestorePoint) {
    Log "Phase 2 - Create Restore Point" "SECTION"
    $s = "$($CONFIG.WorkDir)\Modules\2-RestorePoint.ps1"
    if (Test-Path $s) { & $s } else { Log "Restore point script not found, skipping." "WARN" }
}

# ---------------------------------------------
#  PHASE 3 - HARDWARE REPORT
# ---------------------------------------------
if (-not $CONFIG.SkipHWReport) {
    Log "Phase 3 - Hardware Diagnostics" "SECTION"
    $s = "$($CONFIG.WorkDir)\Modules\3-HardwareReport.ps1"
    if (Test-Path $s) {
        & $s -EmailTo    $CONFIG.ReportEmail `
             -SMTPServer $CONFIG.SMTPServer `
             -SMTPPort   $CONFIG.SMTPPort `
             -SMTPUser   $CONFIG.SMTPUser `
             -SMTPPass   $CONFIG.SMTPPass
    } else { Log "HardwareReport script not found, skipping." "WARN" }
}

# ---------------------------------------------
#  PHASE 4 - INSTALL M365
# ---------------------------------------------
Log "Phase 4 - Microsoft 365 Installation" "SECTION"
$s = "$($CONFIG.WorkDir)\Modules\4-InstallM365.ps1"
if (Test-Path $s) {
    & $s -WorkDir $CONFIG.WorkDir
} else {
    Log "M365 install script not found." "ERROR"
    exit 1
}

# ---------------------------------------------
#  PHASE 5 - CONFIGURE M365
# ---------------------------------------------
Log "Phase 5 - Microsoft 365 Account Configuration" "SECTION"
$s = "$($CONFIG.WorkDir)\Modules\5-ConfigureM365.ps1"
if (Test-Path $s) {
    & $s -UserPrincipalName $M365User -Password $M365Pass
} else {
    Log "M365 config script not found." "WARN"
}

# ---------------------------------------------
#  CLEAR SENSITIVE DATA FROM MEMORY
# ---------------------------------------------
$smtpPlain    = $null
$CONFIG.SMTPPass = $null
$M365Pass     = $null

# ---------------------------------------------
#  DONE
# ---------------------------------------------
Log "All phases complete" "SECTION"
Log "Setup finished successfully. Log saved to: $($CONFIG.LogFile)" "OK"
Write-Host ""
Write-Host "  +==================================================+" -ForegroundColor Green
Write-Host "  |   [OK]  Setup Complete!                             |" -ForegroundColor Green
Write-Host "  |      Microsoft 365 is installed & configured.    |" -ForegroundColor Green
Write-Host "  +==================================================+" -ForegroundColor Green
Write-Host ""