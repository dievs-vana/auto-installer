#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Module 1 - Windows Debloat
    Removes common pre-installed bloatware and disables unnecessary services/telemetry.
#>

param()

function Log {
    param([string]$Msg, [string]$Level = "INFO")
    $ts = Get-Date -Format "HH:mm:ss"
    switch ($Level) {
        "OK"    { Write-Host "    [[OK]] $Msg" -ForegroundColor Green }
        "WARN"  { Write-Host "    [!] $Msg" -ForegroundColor Yellow }
        "ERROR" { Write-Host "    [[FAIL]] $Msg" -ForegroundColor Red }
        default { Write-Host "    [..] $Msg" -ForegroundColor White }
    }
}

# -- Bloatware App List ------------------------------------------------------
$BloatApps = @(
    # Microsoft bloat
    "Microsoft.3DBuilder"
    "Microsoft.BingWeather"
    "Microsoft.BingNews"
    "Microsoft.BingFinance"
    "Microsoft.BingSports"
    "Microsoft.BingTravel"
    "Microsoft.BingFoodAndDrink"
    "Microsoft.BingHealthAndFitness"
    "Microsoft.GetHelp"
    "Microsoft.Getstarted"
    "Microsoft.MicrosoftSolitaireCollection"
    "Microsoft.MixedReality.Portal"
    "Microsoft.People"
    "Microsoft.SkypeApp"
    "Microsoft.Todos"
    "Microsoft.WindowsAlarms"
    "Microsoft.WindowsFeedbackHub"
    "Microsoft.WindowsMaps"
    "Microsoft.WindowsSoundRecorder"
    "Microsoft.Xbox.TCUI"
    "Microsoft.XboxApp"
    "Microsoft.XboxGameOverlay"
    "Microsoft.XboxGamingOverlay"
    "Microsoft.XboxIdentityProvider"
    "Microsoft.XboxSpeechToTextOverlay"
    "Microsoft.YourPhone"
    "Microsoft.ZuneMusic"
    "Microsoft.ZuneVideo"
    "Microsoft.GamingApp"
    "MicrosoftTeams"
    # Third-party
    "king.com.CandyCrushSaga"
    "king.com.CandyCrushFriends"
    "king.com.BubbleWitch3Saga"
    "Facebook.Facebook"
    "Spotify.Spotify"
    "TikTok.TikTok"
    "BytedancePte.Ltd.TikTok"
    "9E2F88E3.Twitter"
    "AdobeSystemsIncorporated.AdobePhotoshopExpress"
    "PandoraMediaInc.29680B314EFC2"
    "Clipchamp.Clipchamp"
)

# -- Remove Bloatware --------------------------------------------------------
Log "Removing bloatware apps..."
foreach ($app in $BloatApps) {
    $pkg = Get-AppxPackage -Name $app -AllUsers -ErrorAction SilentlyContinue
    if ($pkg) {
        try {
            Remove-AppxPackage -Package $pkg.PackageFullName -AllUsers -ErrorAction SilentlyContinue
            # Also remove provisioned so it won't reinstall for new users
            Get-AppxProvisionedPackage -Online |
                Where-Object DisplayName -EQ $app |
                Remove-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue | Out-Null
            Log "Removed: $app" "OK"
        } catch {
            Log "Could not remove: $app" "WARN"
        }
    }
}

# -- Disable Telemetry -------------------------------------------------------
Log "Disabling telemetry and data collection..."

$telemetryKeys = @{
    "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection"     = @{ AllowTelemetry = 0 }
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\DataCollection" = @{
        AllowTelemetry        = 0
        MaxTelemetryAllowed   = 0
    }
}

foreach ($path in $telemetryKeys.Keys) {
    if (-not (Test-Path $path)) { New-Item -Path $path -Force | Out-Null }
    foreach ($name in $telemetryKeys[$path].Keys) {
        Set-ItemProperty -Path $path -Name $name -Value $telemetryKeys[$path][$name] -Type DWord -Force
    }
}
Log "Telemetry disabled." "OK"

# -- Disable Cortana ---------------------------------------------------------
Log "Disabling Cortana..."
$cortanaPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search"
if (-not (Test-Path $cortanaPath)) { New-Item -Path $cortanaPath -Force | Out-Null }
Set-ItemProperty -Path $cortanaPath -Name "AllowCortana" -Value 0 -Type DWord -Force
Log "Cortana disabled." "OK"

# -- Disable Consumer Features (reinstall of bloat from Store) ---------------
Log "Disabling Windows consumer features..."
$consumerPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent"
if (-not (Test-Path $consumerPath)) { New-Item -Path $consumerPath -Force | Out-Null }
Set-ItemProperty -Path $consumerPath -Name "DisableWindowsConsumerFeatures" -Value 1 -Type DWord -Force
Log "Consumer features disabled." "OK"

# -- Disable OneDrive Auto-Start (not uninstall - M365 may use it) -----------
Log "Disabling OneDrive auto-start (can be re-enabled after M365 login)..."
$oneDriveRun = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
Remove-ItemProperty -Path $oneDriveRun -Name "OneDrive" -ErrorAction SilentlyContinue
Log "OneDrive auto-start disabled." "OK"

# -- Disable Unnecessary Scheduled Tasks -------------------------------------
$tasksToDisable = @(
    "\Microsoft\Windows\Application Experience\Microsoft Compatibility Appraiser"
    "\Microsoft\Windows\Application Experience\ProgramDataUpdater"
    "\Microsoft\Windows\Autochk\Proxy"
    "\Microsoft\Windows\Customer Experience Improvement Program\Consolidator"
    "\Microsoft\Windows\Customer Experience Improvement Program\UsbCeip"
    "\Microsoft\Windows\DiskDiagnostic\Microsoft-Windows-DiskDiagnosticDataCollector"
)

Log "Disabling unnecessary scheduled tasks..."
foreach ($task in $tasksToDisable) {
    try {
        Disable-ScheduledTask -TaskPath (Split-Path $task) -TaskName (Split-Path $task -Leaf) -ErrorAction SilentlyContinue | Out-Null
        Log "Disabled task: $task" "OK"
    } catch {
        Log "Task not found (OK): $task" "WARN"
    }
}

# -- Disable Advertising ID ---------------------------------------------------
Log "Disabling Advertising ID..."
$adPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\AdvertisingInfo"
if (-not (Test-Path $adPath)) { New-Item -Path $adPath -Force | Out-Null }
Set-ItemProperty -Path $adPath -Name "Enabled" -Value 0 -Type DWord -Force
Log "Advertising ID disabled." "OK"

Log "Debloat complete." "OK"