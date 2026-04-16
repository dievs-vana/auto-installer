#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Module 2 — Create System Restore Point
#>

param()

function Log {
    param([string]$Msg, [string]$Level = "INFO")
    switch ($Level) {
        "OK"    { Write-Host "    [OK] $Msg" -ForegroundColor Green }
        "WARN"  { Write-Host "    [!!] $Msg" -ForegroundColor Yellow }
        "ERROR" { Write-Host "    [FAIL] $Msg" -ForegroundColor Red }
        default { Write-Host "    [..] $Msg" -ForegroundColor White }
    }
}

Log "Enabling System Restore on C: drive..."
try {
    Enable-ComputerRestore -Drive "C:\" -ErrorAction Stop
    Log "System Restore enabled on C:\" "OK"
} catch {
    Log "System Restore may already be enabled: $_" "WARN"
}

# Ensure VSS service is running
Log "Starting Volume Shadow Copy service..."
Set-Service -Name VSS -StartupType Manual -ErrorAction SilentlyContinue
Start-Service -Name VSS -ErrorAction SilentlyContinue
Log "VSS service started." "OK"

# Set restore point frequency override (Windows throttles to 1 per 24h by default)
$rpFreqPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore"
Set-ItemProperty -Path $rpFreqPath -Name "SystemRestorePointCreationFrequency" -Value 0 -Type DWord -Force

# Create the restore point
$description = "Pre-M365-Install $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
Log "Creating restore point: '$description'..."

try {
    Checkpoint-Computer -Description $description -RestorePointType "MODIFY_SETTINGS" -ErrorAction Stop
    Log "Restore point created successfully." "OK"
} catch {
    Log "Failed to create restore point: $_" "ERROR"
    Log "Continuing setup anyway..." "WARN"
}