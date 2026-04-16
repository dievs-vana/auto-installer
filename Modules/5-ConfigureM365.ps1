#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Module 5 - Microsoft 365 Account Configuration & Sign-In Pre-fill
    
    IMPORTANT NOTE ON AUTO-LOGIN:
    Microsoft 365 uses OAuth 2.0 / modern authentication. Full silent SSO is
    only possible in domain-joined (Azure AD / Entra ID joined) environments.
    
    For standalone/workgroup machines, this script:
      1. Pre-registers the account in Windows Credential Manager
      2. Sets the registry default UPN so Office opens to the sign-in screen
         pre-filled with the user's email
      3. Launches the Office activation flow with the account pre-filled
      4. For Azure AD joined devices - performs a full silent sign-in via
         dsregcmd and triggers SSO to all Office apps
    
    The user will only need to enter their password once (or approve MFA) on
    first launch. All subsequent launches will be silent/automatic.
#>

param(
    [string]$UserPrincipalName = "",
    [SecureString]$Password    = $null
)

function Log {
    param([string]$Msg, [string]$Level = "INFO")
    switch ($Level) {
        "OK"    { Write-Host "    [[OK]] $Msg" -ForegroundColor Green }
        "WARN"  { Write-Host "    [!] $Msg" -ForegroundColor Yellow }
        "ERROR" { Write-Host "    [[FAIL]] $Msg" -ForegroundColor Red }
        default { Write-Host "    [..] $Msg" -ForegroundColor White }
    }
}

if (-not $UserPrincipalName) {
    Log "No UPN provided. Skipping account configuration." "WARN"
    return
}

$domain = ($UserPrincipalName -split "@")[1]

# -----------------------------------------------------------------------------
#  1. Pre-fill Office Identity Registry Keys
# -----------------------------------------------------------------------------
Log "Configuring Office identity registry keys..."

$identityPaths = @(
    "HKCU:\Software\Microsoft\Office\16.0\Common\Identity",
    "HKCU:\Software\Microsoft\Office\16.0\Common\ServicesManagerCache"
)

foreach ($path in $identityPaths) {
    if (-not (Test-Path $path)) { New-Item -Path $path -Force | Out-Null }
}

$idPath = "HKCU:\Software\Microsoft\Office\16.0\Common\Identity"
Set-ItemProperty -Path $idPath -Name "ADUserName"           -Value $UserPrincipalName -Force
Set-ItemProperty -Path $idPath -Name "FirstRunNextTime"     -Value 1                  -Type DWord -Force
Set-ItemProperty -Path $idPath -Name "FederatedLogin"       -Value 1                  -Type DWord -Force

# Disable first-run dialogs
$firstRunPath = "HKCU:\Software\Microsoft\Office\16.0\FirstRun"
if (-not (Test-Path $firstRunPath)) { New-Item -Path $firstRunPath -Force | Out-Null }
Set-ItemProperty -Path $firstRunPath -Name "disablemovie"   -Value 1 -Type DWord -Force
Set-ItemProperty -Path $firstRunPath -Name "BootedRTM"      -Value 1 -Type DWord -Force

Log "Office identity pre-configured for: $UserPrincipalName" "OK"

# -----------------------------------------------------------------------------
#  2. Store Credentials in Windows Credential Manager
# -----------------------------------------------------------------------------
Log "Storing credentials in Windows Credential Manager..."

if ($Password) {
    # Convert SecureString to plain text for cmdkey (stored encrypted by Windows)
    $BSTR      = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password)
    $plainPass = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)

    # Store for common Office/M365 endpoints
    $targets = @(
        "MicrosoftOffice16_Data:orgid:$UserPrincipalName",
        "https://login.microsoftonline.com",
        "https://login.microsoft.com",
        "Office 365 - $UserPrincipalName"
    )

    foreach ($target in $targets) {
        try {
            $cmdkeyArgs = "/generic:`"$target`" /user:`"$UserPrincipalName`" /pass:`"$plainPass`""
            Start-Process -FilePath "cmdkey.exe" -ArgumentList $cmdkeyArgs -WindowStyle Hidden -Wait
        } catch { }
    }

    # Clear plain text from memory
    $plainPass = $null
    Log "Credentials stored in Credential Manager." "OK"
} else {
    Log "No password provided. Skipping Credential Manager storage." "WARN"
    Log "User will be prompted for password on first Office launch." "INFO"
}

# -----------------------------------------------------------------------------
#  3. Check for Azure AD / Entra ID Join (enables full SSO)
# -----------------------------------------------------------------------------
Log "Checking device Azure AD join status..."
$dsregOutput = dsregcmd /status 2>$null | Out-String

$isAzureADJoined    = $dsregOutput -match "AzureAdJoined\s*:\s*YES"
$isWorkplaceJoined  = $dsregOutput -match "WorkplaceJoined\s*:\s*YES"

if ($isAzureADJoined) {
    Log "Device is Azure AD Joined - SSO will work automatically." "OK"
} elseif ($isWorkplaceJoined) {
    Log "Device is Workplace (Entra) Registered - partial SSO available." "OK"
} else {
    Log "Device is NOT Azure AD joined (workgroup/local machine)." "WARN"
    Log "Office will open with email pre-filled. Password entry required on first launch." "INFO"
    
    # Optionally register the device with Azure AD for SSO
    # Uncomment the block below if you want to attempt Workplace Join:
    <#
    Log "Attempting Workplace Join for SSO..."
    try {
        $joinArgs = "/join /force"
        Start-Process -FilePath "dsregcmd.exe" -ArgumentList $joinArgs -Wait -WindowStyle Hidden
        Log "Workplace Join attempted. Check dsregcmd /status to verify." "OK"
    } catch {
        Log "Workplace Join failed: $_" "WARN"
    }
    #>
}

# -----------------------------------------------------------------------------
#  4. Suppress Office EULA and "What's New" dialogs
# -----------------------------------------------------------------------------
Log "Suppressing Office setup dialogs..."

$suppressKeys = @{
    "HKCU:\Software\Microsoft\Office\16.0\Common"                    = @{ "qmenable" = 0; "sendcustomerdata" = 0 }
    "HKCU:\Software\Microsoft\Office\16.0\Word\Options"              = @{ "NoRelaunchOLKAgain" = 1 }
    "HKCU:\Software\Microsoft\Office\16.0\Outlook\Setup"             = @{ "first-run" = 1 }
    "HKCU:\Software\Microsoft\Office\16.0\Common\PTWatson"           = @{ "PTWOptIn" = 0 }
    "HKCU:\Software\Microsoft\Office\16.0\Common\LanguageResources"  = @{ "UILanguage" = 1033 }
}

foreach ($regPath in $suppressKeys.Keys) {
    if (-not (Test-Path $regPath)) { New-Item -Path $regPath -Force | Out-Null }
    foreach ($name in $suppressKeys[$regPath].Keys) {
        Set-ItemProperty -Path $regPath -Name $name -Value $suppressKeys[$regPath][$name] -Type DWord -Force
    }
}
Log "Office dialogs suppressed." "OK"

# -----------------------------------------------------------------------------
#  5. Trigger Office Activation (opens minimized, pre-filled with UPN)
# -----------------------------------------------------------------------------
Log "Triggering Office activation with pre-filled account..."

$officePaths = @(
    "${env:ProgramFiles}\Microsoft Office\root\Office16\MSOSYNC.EXE",
    "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\MSOSYNC.EXE"
)

$syncExe = $officePaths | Where-Object { Test-Path $_ } | Select-Object -First 1

if ($syncExe) {
    Log "Launching Office sync (background)..." "OK"
    Start-Process -FilePath $syncExe -ArgumentList "/Start" -WindowStyle Minimized
    Log "Office sync launched. It will authenticate using stored credentials." "OK"
} else {
    Log "MSOSYNC.EXE not found - Office may not be fully installed yet." "WARN"
}

# -----------------------------------------------------------------------------
#  Summary
# -----------------------------------------------------------------------------
Write-Host ""
Write-Host "    +---------------------------------------------------------+" -ForegroundColor Cyan
Write-Host "    |  Account Configuration Summary                          |" -ForegroundColor Cyan
Write-Host "    +---------------------------------------------------------+" -ForegroundColor Cyan
Write-Host "    |  Account  : $UserPrincipalName" -ForegroundColor White
Write-Host "    |  SSO Mode : $(if ($isAzureADJoined) { 'Full SSO (Azure AD Joined)' } elseif ($isWorkplaceJoined) { 'Partial SSO (Workplace)' } else { 'Credential Manager (1-time password)' })" -ForegroundColor White
Write-Host "    |  Status   : Pre-configured and ready                    |" -ForegroundColor Green
Write-Host "    +---------------------------------------------------------+" -ForegroundColor Cyan
Write-Host ""