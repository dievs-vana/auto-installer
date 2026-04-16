#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Module 4 — Silent Microsoft 365 Installation
    Uses the Office Deployment Tool (ODT) for a completely silent install.
#>

param(
    [string]$WorkDir = "$env:TEMP\M365Setup"
)

function Log {
    param([string]$Msg, [string]$Level = "INFO")
    switch ($Level) {
        "OK"    { Write-Host "    [✓] $Msg" -ForegroundColor Green }
        "WARN"  { Write-Host "    [!] $Msg" -ForegroundColor Yellow }
        "ERROR" { Write-Host "    [✗] $Msg" -ForegroundColor Red }
        default { Write-Host "    [•] $Msg" -ForegroundColor White }
    }
}

# ── Paths ──────────────────────────────────────────────────────────────────────
$ODTDir     = Join-Path $WorkDir "ODT"
$ODTExe     = Join-Path $ODTDir  "setup.exe"
$ConfigXml  = Join-Path $WorkDir "Config\ODT-Config.xml"
$ODTUrl     = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_17328-20162.exe"

New-Item -ItemType Directory -Force -Path $ODTDir | Out-Null

# ── Check if Office / M365 is already installed ───────────────────────────────
function Test-OfficeInstalled {
    <#
    .SYNOPSIS
        Returns a [PSCustomObject] with:
          Found   [bool]   — whether any Office install was detected
          Version [string] — version string if available, else ""
          Method  [string] — which detection method matched
        
        Detection order (first match wins and returns immediately):
          1. Click-to-Run registry  — modern M365 / Office 2019/2021/2024
          2. MSI registry           — legacy volume/retail Office installs
          3. Add/Remove Programs    — catches any installer-registered Office
          4. Executable on disk     — final safety net (WINWORD.EXE)
    #>

    # 1. Click-to-Run (modern Office / M365) ───────────────────────────────────
    $c2rKeys = @(
        "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration"
    )
    foreach ($key in $c2rKeys) {
        if (Test-Path $key) {
            $props = Get-ItemProperty $key -ErrorAction SilentlyContinue
            if ($props.VersionToReport) {
                return [PSCustomObject]@{ Found = $true; Version = $props.VersionToReport; Method = "Click-to-Run registry" }
            }
        }
    }

    # 2. MSI-based Office (legacy volume / retail) ─────────────────────────────
    $msiKeys = @(
        "HKLM:\SOFTWARE\Microsoft\Office\16.0\Common\InstallRoot",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\16.0\Common\InstallRoot",
        "HKLM:\SOFTWARE\Microsoft\Office\15.0\Common\InstallRoot",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\Common\InstallRoot"
    )
    foreach ($key in $msiKeys) {
        if (Test-Path $key) {
            $path = (Get-ItemProperty $key -ErrorAction SilentlyContinue).Path
            if ($path) {
                $ver = if ($key -match "16\.0") { "16.x (Office 2016/2019/2021 MSI)" } else { "15.x (Office 2013 MSI)" }
                return [PSCustomObject]@{ Found = $true; Version = $ver; Method = "MSI registry" }
            }
        }
    }

    # 3. Add/Remove Programs (covers any installer-registered Office build) ────
    $arpPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    foreach ($arpPath in $arpPaths) {
        $entry = Get-ItemProperty $arpPath -ErrorAction SilentlyContinue |
                 Where-Object {
                     $_.DisplayName -match "Microsoft 365|Microsoft Office" -and
                     $_.DisplayName -notmatch "Teams|OneDrive|OneNote standalone"
                 } | Select-Object -First 1
        if ($entry) {
            $ver = if ($entry.DisplayVersion) { $entry.DisplayVersion } else { "unknown" }
            return [PSCustomObject]@{ Found = $true; Version = $ver; Method = "Add/Remove Programs ($($entry.DisplayName))" }
        }
    }

    # 4. Executable on disk (final safety net) ─────────────────────────────────
    $exePaths = @(
        "${env:ProgramFiles}\Microsoft Office\root\Office16\WINWORD.EXE",
        "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\WINWORD.EXE",
        "${env:ProgramFiles}\Microsoft Office\Office16\WINWORD.EXE",
        "${env:ProgramFiles(x86)}\Microsoft Office\Office16\WINWORD.EXE"
    )
    foreach ($exe in $exePaths) {
        if (Test-Path $exe) {
            $fileVer = (Get-Item $exe -ErrorAction SilentlyContinue).VersionInfo.FileVersion
            $ver = if ($fileVer) { $fileVer } else { "unknown" }
            return [PSCustomObject]@{ Found = $true; Version = $ver; Method = "Executable on disk ($exe)" }
        }
    }

    return [PSCustomObject]@{ Found = $false; Version = ""; Method = "" }
}

Log "Checking for existing Office / Microsoft 365 installation..."
$detection = Test-OfficeInstalled

if ($detection.Found) {
    Log "Microsoft 365 / Office is already installed — skipping installation." "WARN"
    Log "Detected via : $($detection.Method)" "WARN"
    Log "Version      : $($detection.Version)" "WARN"
    Log "If you want to reinstall or upgrade, remove Office first via Settings → Apps." "WARN"
    return
}

# ── Download ODT ──────────────────────────────────────────────────────────────
Log "Downloading Office Deployment Tool..."
$odtInstaller = Join-Path $ODTDir "odt_installer.exe"
try {
    Invoke-WebRequest -Uri $ODTUrl -OutFile $odtInstaller -UseBasicParsing
    Log "ODT downloaded." "OK"
} catch {
    Log "Failed to download ODT: $_" "ERROR"
    Log "Please download manually from: https://www.microsoft.com/en-us/download/details.aspx?id=49117" "WARN"
    exit 1
}

# ── Extract ODT ──────────────────────────────────────────────────────────────
Log "Extracting ODT..."
$extractArgs = "/quiet /extract:`"$ODTDir`""
Start-Process -FilePath $odtInstaller -ArgumentList $extractArgs -Wait -NoNewWindow
if (-not (Test-Path $ODTExe)) {
    Log "ODT setup.exe not found after extraction. Checking directory..." "ERROR"
    Get-ChildItem $ODTDir | ForEach-Object { Log "  Found: $($_.Name)" "WARN" }
    exit 1
}
Log "ODT extracted." "OK"

# ── Use config from repo or fallback inline ───────────────────────────────────
if (-not (Test-Path $ConfigXml)) {
    Log "ODT-Config.xml not found in repo. Using built-in default config..." "WARN"
    $ConfigXml = Join-Path $ODTDir "ODT-Config.xml"
    @"
<Configuration ID="M365Setup">
  <Add OfficeClientEdition="64" Channel="Current">
    <Product ID="O365BusinessRetail">
      <Language ID="en-us" />
      <ExcludeApp ID="Access" />
      <ExcludeApp ID="Groove" />
      <ExcludeApp ID="Lync" />
      <ExcludeApp ID="Publisher" />
    </Product>
  </Add>
  <Updates Enabled="TRUE" Channel="Current" />
  <Display Level="None" AcceptEULA="TRUE" />
  <Logging Level="Standard" Path="%temp%\M365Setup\ODT-Logs" />
  <Property Name="AUTOACTIVATE" Value="1" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
  <Property Name="SharedComputerLicensing" Value="0" />
  <Property Name="PinIconsToTaskbar" Value="TRUE" />
  <RemoveMSI />
</Configuration>
"@ | Set-Content -Path $ConfigXml -Encoding UTF8
    Log "Built-in config written to: $ConfigXml" "OK"
}

# ── Download Office Files ─────────────────────────────────────────────────────
Log "Downloading Microsoft 365 installation files (this may take 5-15 minutes)..."
Log "Please wait — downloading ~2-3 GB..." "WARN"

$downloadArgs = "/download `"$ConfigXml`""
$dlProcess = Start-Process -FilePath $ODTExe -ArgumentList $downloadArgs -Wait -NoNewWindow -PassThru
if ($dlProcess.ExitCode -ne 0) {
    Log "ODT download phase exited with code: $($dlProcess.ExitCode)" "WARN"
    Log "Attempting install directly (online mode)..." "WARN"
}

# ── Silent Install ────────────────────────────────────────────────────────────
Log "Starting silent Microsoft 365 installation..."
$installArgs = "/configure `"$ConfigXml`""
$installProcess = Start-Process -FilePath $ODTExe -ArgumentList $installArgs -Wait -NoNewWindow -PassThru

if ($installProcess.ExitCode -eq 0) {
    Log "Microsoft 365 installed successfully!" "OK"
} elseif ($installProcess.ExitCode -eq 3010) {
    Log "Microsoft 365 installed — a reboot is required to complete." "WARN"
} else {
    Log "Installation exited with code: $($installProcess.ExitCode)" "ERROR"
    Log "Check ODT logs at: $env:TEMP\M365Setup\ODT-Logs" "WARN"
    exit 1
}

# ── Verify Installation ───────────────────────────────────────────────────────
Log "Verifying installation..."
Start-Sleep -Seconds 5
$wordPath   = "${env:ProgramFiles}\Microsoft Office\root\Office16\WINWORD.EXE"
$wordPath86 = "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\WINWORD.EXE"

if ((Test-Path $wordPath) -or (Test-Path $wordPath86)) {
    Log "Microsoft 365 verified — Word.exe found." "OK"
} else {
    Log "Verification check: Word.exe not found at expected path. Office may still be installed in a non-standard location." "WARN"
}