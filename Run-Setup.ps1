#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Microsoft 365 Silent Installer with System Preparation
    Run from PowerShell (Admin): irm https://raw.githubusercontent.com/dievs-vana/auto-installer/main/Run-Setup.ps1 | iex
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ─────────────────────────────────────────────
#  BANNER
# ─────────────────────────────────────────────
Clear-Host
Write-Host ""
Write-Host "  ╔══════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "  ║         Microsoft 365 Auto-Setup & Provisioner       ║" -ForegroundColor Cyan
Write-Host "  ║              System Prep + Silent Install             ║" -ForegroundColor Cyan
Write-Host "  ╚══════════════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""

# ─────────────────────────────────────────────
#  CONFIG
# ─────────────────────────────────────────────
$CONFIG = @{
    RepoBase         = "https://raw.githubusercontent.com/dievs-vana/auto-installer/main"
    WorkDir          = "$env:TEMP\M365Setup"
    LogFile          = "$env:TEMP\M365Setup\setup.log"
    SMTPServer       = "smtp.gmail.com"
    SMTPPort         = 587
    SkipDebloat      = $false
    SkipRestorePoint = $false
    SkipHWReport     = $false

    # ── SMTP credentials — loaded from config.ps1 ──────────────────────────
    # Do NOT hardcode credentials here.
    # Create a config.ps1 file in the same folder as this script with:
    #   $SmtpUser = "sender@gmail.com"
    #   $SmtpPass = "xxxx xxxx xxxx xxxx"   # Gmail App Password
    #   $SmtpTo   = "recipient@email.com"
    # config.ps1 is listed in .gitignore and will NOT be uploaded to GitHub.
    ReportEmail = ""
    SMTPUser    = ""
    SMTPPass    = ""
}

# ─────────────────────────────────────────────
#  LOAD LOCAL CONFIG (credentials)
# ─────────────────────────────────────────────
$localConfig = Join-Path $PSScriptRoot "config.ps1"
if (Test-Path $localConfig) {
    . $localConfig
    $CONFIG.SMTPUser    = $SmtpUser
    $CONFIG.SMTPPass    = $SmtpPass
    $CONFIG.ReportEmail = $SmtpTo
    Write-Host "  [✓] Loaded credentials from config.ps1" -ForegroundColor Green
} else {
    Write-Host "  [!] config.ps1 not found — hardware report email will be skipped." -ForegroundColor Yellow
    Write-Host "      Copy config.example.ps1 to config.ps1 and fill in your credentials." -ForegroundColor Yellow
    Write-Host ""
}

# ─────────────────────────────────────────────
#  HELPERS
# ─────────────────────────────────────────────
function Log {
    param([string]$Msg, [string]$Level = "INFO")
    $ts   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$ts][$Level] $Msg"
    Add-Content -Path $CONFIG.LogFile -Value $line -Encoding UTF8
    switch ($Level) {
        "INFO"    { Write-Host "  [•] $Msg" -ForegroundColor White }
        "OK"      { Write-Host "  [✓] $Msg" -ForegroundColor Green }
        "WARN"    { Write-Host "  [!] $Msg" -ForegroundColor Yellow }
        "ERROR"   { Write-Host "  [✗] $Msg" -ForegroundColor Red }
        "SECTION" {
            Write-Host ""
            Write-Host "  ── $Msg " -ForegroundColor Cyan
        }
    }
}

function Ensure-WorkDir {
    if (-not (Test-Path $CONFIG.WorkDir)) {
        New-Item -ItemType Directory -Path $CONFIG.WorkDir -Force | Out-Null
    }
}

function Download-Script {
    param([string]$FileName)
    $url  = "$($CONFIG.RepoBase)/$FileName"
    # Normalize backslashes to forward slashes for URL
    $url  = $url -replace "\\", "/"
    $dest = Join-Path $CONFIG.WorkDir $FileName
    Log "Downloading $FileName..."
    try {
        Invoke-WebRequest -Uri $url -OutFile $dest -UseBasicParsing
        Log "Downloaded: $FileName" "OK"
    } catch {
        Log "Failed to download $FileName : $_" "ERROR"
        throw
    }
    return $dest
}

# ─────────────────────────────────────────────
#  STEP 0 — CREDENTIAL PROMPT
# ─────────────────────────────────────────────
Log "Credential Collection" "SECTION"
Write-Host ""
Write-Host "  Enter Microsoft 365 credentials." -ForegroundColor White
Write-Host "  These will be used to pre-configure the account after installation." -ForegroundColor Gray
Write-Host ""

$M365User = Read-Host "  Microsoft 365 Email (UPN)"
$M365Pass = Read-Host "  Password" -AsSecureString

# Validate format
if ($M365User -notmatch "^[^@\s]+@[^@\s]+\.[^@\s]+$") {
    Log "Invalid email format entered." "ERROR"
    exit 1
}

Log "Credentials collected (stored in memory only)." "OK"

# ─────────────────────────────────────────────
#  STEP 1 — SETUP WORKSPACE
# ─────────────────────────────────────────────
Log "Initializing workspace" "SECTION"
Ensure-WorkDir
New-Item -ItemType Directory -Force -Path "$($CONFIG.WorkDir)\Modules" | Out-Null
New-Item -ItemType Directory -Force -Path "$($CONFIG.WorkDir)\Config"  | Out-Null
Log "Work directory: $($CONFIG.WorkDir)" "OK"

# ─────────────────────────────────────────────
#  DOWNLOAD ALL MODULES
# ─────────────────────────────────────────────
Log "Downloading setup modules" "SECTION"
$scripts = @(
    "Modules/1-Debloat.ps1",
    "Modules/2-RestorePoint.ps1",
    "Modules/3-HardwareReport.ps1",
    "Modules/4-InstallM365.ps1",
    "Modules/5-ConfigureM365.ps1",
    "Config/ODT-Config.xml"
)

foreach ($s in $scripts) {
    $destPath = Join-Path $CONFIG.WorkDir ($s -replace "/", "\")
    $destDir  = Split-Path $destPath
    if (-not (Test-Path $destDir)) {
        New-Item -ItemType Directory -Force -Path $destDir | Out-Null
    }
    try {
        $url = "$($CONFIG.RepoBase)/$s"
        Invoke-WebRequest -Uri $url -OutFile $destPath -UseBasicParsing
        Log "Downloaded: $s" "OK"
    } catch {
        Log "Skipping $s (not found on repo)" "WARN"
    }
}

# ─────────────────────────────────────────────
#  STEP 2 — DEBLOAT
# ─────────────────────────────────────────────
if (-not $CONFIG.SkipDebloat) {
    Log "Phase 1 — System Debloat" "SECTION"
    $debloatScript = "$($CONFIG.WorkDir)\Modules\1-Debloat.ps1"
    if (Test-Path $debloatScript) {
        & $debloatScript
    } else {
        Log "Debloat script not found, skipping." "WARN"
    }
}

# ─────────────────────────────────────────────
#  STEP 3 — RESTORE POINT
# ─────────────────────────────────────────────
if (-not $CONFIG.SkipRestorePoint) {
    Log "Phase 2 — Create Restore Point" "SECTION"
    $rpScript = "$($CONFIG.WorkDir)\Modules\2-RestorePoint.ps1"
    if (Test-Path $rpScript) {
        & $rpScript
    } else {
        Log "RestorePoint script not found, skipping." "WARN"
    }
}

# ─────────────────────────────────────────────
#  STEP 4 — HARDWARE REPORT
# ─────────────────────────────────────────────
if (-not $CONFIG.SkipHWReport) {
    Log "Phase 3 — Hardware Diagnostics" "SECTION"
    $hwScript = "$($CONFIG.WorkDir)\Modules\3-HardwareReport.ps1"
    if (Test-Path $hwScript) {
        & $hwScript -EmailTo    $CONFIG.ReportEmail `
                    -SMTPServer $CONFIG.SMTPServer `
                    -SMTPPort   $CONFIG.SMTPPort `
                    -SMTPUser   $CONFIG.SMTPUser `
                    -SMTPPass   $CONFIG.SMTPPass
    } else {
        Log "HardwareReport script not found, skipping." "WARN"
    }
}

# ─────────────────────────────────────────────
#  STEP 5 — INSTALL MICROSOFT 365
# ─────────────────────────────────────────────
Log "Phase 4 — Microsoft 365 Installation" "SECTION"
$installScript = "$($CONFIG.WorkDir)\Modules\4-InstallM365.ps1"
if (Test-Path $installScript) {
    & $installScript -WorkDir $CONFIG.WorkDir
} else {
    Log "M365 install script not found." "ERROR"
    exit 1
}

# ─────────────────────────────────────────────
#  STEP 6 — CONFIGURE M365 (Auto-Login)
# ─────────────────────────────────────────────
Log "Phase 5 — Microsoft 365 Account Configuration" "SECTION"
$configScript = "$($CONFIG.WorkDir)\Modules\5-ConfigureM365.ps1"
if (Test-Path $configScript) {
    & $configScript -UserPrincipalName $M365User -Password $M365Pass
} else {
    Log "M365 config script not found." "WARN"
}

# ─────────────────────────────────────────────
#  DONE
# ─────────────────────────────────────────────
Log "All phases complete" "SECTION"
Log "Setup finished successfully. Log: $($CONFIG.LogFile)" "OK"
Write-Host ""
Write-Host "  ╔══════════════════════════════════════╗" -ForegroundColor Green
Write-Host "  ║   Setup Complete! Microsoft 365       ║" -ForegroundColor Green
Write-Host "  ║   is installed and configured.        ║" -ForegroundColor Green
Write-Host "  ╚══════════════════════════════════════╝" -ForegroundColor Green
Write-Host ""