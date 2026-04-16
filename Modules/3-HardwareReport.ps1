#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Module 3 — Hardware Diagnostics Report
    Collects system info and emails an HTML report.
#>

param(
    [string]$EmailTo    = "your@email.com",
    [string]$SMTPServer = "smtp.gmail.com",
    [int]$SMTPPort      = 587,
    [string]$SMTPUser   = "",
    [string]$SMTPPass   = ""
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

Log "Collecting hardware information..."

# ── OS Info ──────────────────────────────────────────────────────────────────
$os      = Get-CimInstance Win32_OperatingSystem
$cs      = Get-CimInstance Win32_ComputerSystem
$cpu     = Get-CimInstance Win32_Processor | Select-Object -First 1
$bios    = Get-CimInstance Win32_BIOS
$board   = Get-CimInstance Win32_BaseBoard

# ── RAM ──────────────────────────────────────────────────────────────────────
$ramModules = Get-CimInstance Win32_PhysicalMemory
$totalRAM_GB = [math]::Round(($ramModules | Measure-Object -Property Capacity -Sum).Sum / 1GB, 2)
$ramDetails = $ramModules | ForEach-Object {
    "$($_.BankLabel) — $([math]::Round($_.Capacity/1GB,0)) GB @ $($_.Speed) MHz ($($_.Manufacturer))"
}

# ── Disk ──────────────────────────────────────────────────────────────────────
$disks = Get-CimInstance Win32_DiskDrive | ForEach-Object {
    $size = [math]::Round($_.Size / 1GB, 1)
    "$($_.Model) — $size GB ($($_.InterfaceType))"
}

# ── Battery ──────────────────────────────────────────────────────────────────
$batteryStatus = "No battery detected (Desktop)"
$batteryHealth = "N/A"
$battery = Get-CimInstance Win32_Battery -ErrorAction SilentlyContinue
if ($battery) {
    $statusMap = @{
        1 = "Other"; 2 = "Unknown"; 3 = "Fully Charged"; 4 = "Low";
        5 = "Critical"; 6 = "Charging"; 7 = "Charging/High"; 8 = "Charging/Low";
        9 = "Charging/Critical"; 10 = "Undefined"; 11 = "Partially Charged"
    }
    $batteryStatus = $statusMap[[int]$battery.BatteryStatus] + " ($($battery.EstimatedChargeRemaining)%)"
    
    # Run powercfg battery report
    $reportPath = "$env:TEMP\M365Setup\battery-report.html"
    powercfg /batteryreport /output $reportPath 2>$null
    $batteryHealth = if (Test-Path $reportPath) { "Battery report generated (see attachment)" } else { "Unable to generate report" }
}

# ── GPU ──────────────────────────────────────────────────────────────────────
$gpus = Get-CimInstance Win32_VideoController | ForEach-Object {
    $vram = if ($_.AdapterRAM) { "$([math]::Round($_.AdapterRAM/1MB,0)) MB VRAM" } else { "Unknown VRAM" }
    "$($_.Name) — $vram (Driver: $($_.DriverVersion))"
}

# ── Network ──────────────────────────────────────────────────────────────────
$nics = Get-CimInstance Win32_NetworkAdapterConfiguration | 
    Where-Object { $_.IPEnabled -eq $true } | ForEach-Object {
    "$($_.Description) — IP: $($_.IPAddress -join ', ')"
}

# ── Uptime ────────────────────────────────────────────────────────────────────
$uptime = (Get-Date) - $os.LastBootUpTime
$uptimeStr = "$([int]$uptime.TotalDays)d $($uptime.Hours)h $($uptime.Minutes)m"

# ── Disk Health via SMART (basic) ─────────────────────────────────────────────
$diskHealth = Get-PhysicalDisk | ForEach-Object {
    "$($_.FriendlyName) — Health: $($_.HealthStatus) | $($_.OperationalStatus)"
}

# ── Windows Activation ───────────────────────────────────────────────────────
$activation = Get-CimInstance SoftwareLicensingProduct -Filter "ApplicationId='55c92734-d682-4d71-983e-d6ec3f16059f' AND LicenseStatus=1" -ErrorAction SilentlyContinue
$activationStatus = if ($activation) { "✅ Activated" } else { "⚠️ Not Activated / Unable to Detect" }

Log "Data collected. Building report..." "OK"

# ── Pending Reboot Check ──────────────────────────────────────────────────────
$pendingReboot = $false
$rebootKeys = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending",
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired"
)
foreach ($key in $rebootKeys) {
    if (Test-Path $key) { $pendingReboot = $true }
}

# ── Build HTML Report ─────────────────────────────────────────────────────────
$timestamp   = Get-Date -Format "MMMM dd, yyyy — HH:mm:ss"
$hostname    = $env:COMPUTERNAME
$currentUser = $env:USERNAME

function HtmlRow {
    param([string]$Label, [string]$Value, [string]$Status = "")
    $statusBadge = if ($Status) {
        $color = switch ($Status) { "ok" { "#22c55e" } "warn" { "#f59e0b" } "error" { "#ef4444" } default { "#6b7280" } }
        "<span style='background:$color;color:#fff;padding:2px 8px;border-radius:4px;font-size:11px;margin-left:8px;'>$Status</span>"
    } else { "" }
    return "<tr><td style='padding:8px 12px;color:#9ca3af;font-size:13px;width:200px;'>$Label</td><td style='padding:8px 12px;color:#f1f5f9;font-size:13px;'>$Value$statusBadge</td></tr>"
}

function HtmlSection {
    param([string]$Title, [string]$Content)
    return @"
    <div style='margin-bottom:24px;'>
      <div style='font-size:11px;letter-spacing:2px;text-transform:uppercase;color:#6366f1;margin-bottom:8px;font-weight:600;'>$Title</div>
      <table style='width:100%;border-collapse:collapse;background:#1e293b;border-radius:8px;overflow:hidden;'>
        $Content
      </table>
    </div>
"@
}

$ramRowsHtml   = ($ramDetails | ForEach-Object { HtmlRow "Module" $_ }) -join ""
$diskRowsHtml  = ($disks       | ForEach-Object { HtmlRow "Drive" $_ }) -join ""
$gpuRowsHtml   = ($gpus        | ForEach-Object { HtmlRow "GPU" $_ }) -join ""
$nicRowsHtml   = ($nics        | ForEach-Object { HtmlRow "Adapter" $_ }) -join ""
$dhRowsHtml    = ($diskHealth  | ForEach-Object { HtmlRow "Disk" $_ }) -join ""
$pendingStr    = if ($pendingReboot) { "Yes — Reboot required" } else { "No" }
$pendingStatus = if ($pendingReboot) { "warn" } else { "ok" }
$battStatus    = if ($battery) { "ok" } else { "" }

$htmlBody = @"
<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"><title>Hardware Report — $hostname</title></head>
<body style='margin:0;padding:0;background:#0f172a;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;'>
  <div style='max-width:720px;margin:40px auto;padding:0 20px;'>

    <div style='background:linear-gradient(135deg,#4f46e5,#7c3aed);border-radius:12px;padding:32px;margin-bottom:32px;'>
      <div style='font-size:11px;letter-spacing:3px;text-transform:uppercase;color:#c4b5fd;margin-bottom:8px;'>System Diagnostics Report</div>
      <h1 style='margin:0;color:#fff;font-size:28px;font-weight:700;'>$hostname</h1>
      <div style='color:#c4b5fd;margin-top:8px;font-size:13px;'>Generated: $timestamp</div>
      <div style='color:#c4b5fd;font-size:13px;'>User: $currentUser</div>
    </div>

    $(HtmlSection "System" (
        (HtmlRow "Manufacturer"    $cs.Manufacturer) +
        (HtmlRow "Model"           $cs.Model) +
        (HtmlRow "OS"              "$($os.Caption) Build $($os.BuildNumber)") +
        (HtmlRow "Architecture"    $os.OSArchitecture) +
        (HtmlRow "Activation"      $activationStatus) +
        (HtmlRow "Last Boot"       $os.LastBootUpTime.ToString("yyyy-MM-dd HH:mm")) +
        (HtmlRow "Uptime"          $uptimeStr) +
        (HtmlRow "Pending Reboot"  $pendingStr $pendingStatus)
    ))

    $(HtmlSection "Processor" (
        (HtmlRow "CPU"    $cpu.Name) +
        (HtmlRow "Cores"  "$($cpu.NumberOfCores) cores / $($cpu.NumberOfLogicalProcessors) logical") +
        (HtmlRow "Speed"  "$($cpu.MaxClockSpeed) MHz max")
    ))

    $(HtmlSection "Memory — $totalRAM_GB GB Total" $ramRowsHtml)

    $(HtmlSection "Storage" ($diskRowsHtml + $dhRowsHtml))

    $(HtmlSection "Display / GPU" $gpuRowsHtml)

    $(HtmlSection "Battery" (
        (HtmlRow "Status" $batteryStatus $battStatus) +
        (HtmlRow "Report" $batteryHealth)
    ))

    $(HtmlSection "Network" $nicRowsHtml)

    $(HtmlSection "Firmware" (
        (HtmlRow "BIOS Vendor"  $bios.Manufacturer) +
        (HtmlRow "BIOS Version" $bios.SMBIOSBIOSVersion) +
        (HtmlRow "BIOS Date"    $bios.ReleaseDate.ToString("yyyy-MM-dd")) +
        (HtmlRow "Baseboard"    "$($board.Manufacturer) $($board.Product)")
    ))

    <div style='text-align:center;color:#475569;font-size:11px;padding:24px 0;'>
      Auto-generated by M365 Setup Script
    </div>
  </div>
</body>
</html>
"@

# Save report locally
$reportFile = "$env:TEMP\M365Setup\HardwareReport-$hostname.html"
$htmlBody | Set-Content -Path $reportFile -Encoding UTF8
Log "Report saved to: $reportFile" "OK"

# ── Send Email ────────────────────────────────────────────────────────────────
if ($SMTPUser -and $SMTPPass -and $EmailTo) {
    Log "Sending report to $EmailTo..."
    try {
        $secPass    = ConvertTo-SecureString $SMTPPass -AsPlainText -Force
        $credential = New-Object System.Management.Automation.PSCredential($SMTPUser, $secPass)

        $mailParams = @{
            To          = $EmailTo
            From        = $SMTPUser
            Subject     = "Hardware Report — $hostname — $(Get-Date -Format 'yyyy-MM-dd')"
            Body        = $htmlBody
            BodyAsHtml  = $true
            SmtpServer  = $SMTPServer
            Port        = $SMTPPort
            Credential  = $credential
            UseSsl      = $true
        }

        # Attach battery report if exists
        $battReportPath = "$env:TEMP\M365Setup\battery-report.html"
        if (Test-Path $battReportPath) {
            $mailParams.Attachments = $battReportPath
        }

        Send-MailMessage @mailParams
        Log "Report emailed successfully to $EmailTo" "OK"
    } catch {
        Log "Email failed: $_" "WARN"
        Log "Report is still saved locally at: $reportFile" "INFO"
    }
} else {
    Log "SMTP credentials not configured. Report saved locally only." "WARN"
}