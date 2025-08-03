<# 
.SYNOPSIS
    Automated Windows Setup Script with Multi-Tool Downloads, Installation, Activation, and Debloating.

.DESCRIPTION
    This script performs the following steps:
    1. Ensures it is running with Administrator privileges
    2. Temporarily relaxes PowerShell's Execution Policy for compatibility
    3. Downloads necessary installer files for Ninite, Office 365, and WinRAR license key
    4. Runs Ninite and Office 365 installers silently where possible
    5. Automatically debloats Windows 10 using the official Windows10SysPrepDebloater script
    6. Attempts to activate Windows automatically using built-in OEM key or fallback scripts
    7. Copies WinRAR registration key if WinRAR is installed
    8. Cleans up all temporary files created during the process
    9. Presents a summary GUI with the task results
    10. Waits for user acknowledgment before exiting

.NOTES
    - Run this script as Administrator to work correctly.
    - Review the URLs used for downloads to ensure trusted sources.
    - Windows Activation steps may not work depending on your license type.
    - Use of debloat and activation scripts can have system impacts; test in a safe environment.

#>

# --- Function: Ensures script runs with Administrator privileges ---
function Ensure-Admin {
    # Check if current user has Administrator role
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent())
                .IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (-not $isAdmin) {
        Write-Warning "This script requires Administrator privileges."
        Write-Warning "Restarting script as administrator..."
        # Relaunch PowerShell running as admin with the same script file path
        Start-Process powershell.exe "-ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
        exit  # Exit current script so relaunch takes over
    }
}

# --- Function: Temporarily sets Execution Policy to Bypass if needed ---
function Ensure-ExecutionPolicy {
    $policy = Get-ExecutionPolicy
    if ($policy -notin @("Bypass", "Unrestricted")) {
        Write-Host "Setting Execution Policy to Bypass for this process..."
        Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
    } else {
        Write-Host "Execution Policy is already permissive: $policy"
    }
}

# --- Function: Downloads a file from a URL safely with retry ---
function Download-FileSafe {
    param (
        [Parameter(Mandatory=$true)][string]$url,
        [Parameter(Mandatory=$true)][string]$outputPath,
        [int]$timeoutSeconds = 15
    )
    Write-Host "Downloading: $url" -ForegroundColor Cyan
    try {
        Invoke-WebRequest -Uri $url -OutFile $outputPath -TimeoutSec $timeoutSeconds -UseBasicParsing -ErrorAction Stop
        Write-Host "Download succeeded: $outputPath" -ForegroundColor Green
        return $true
    } catch {
        Write-Warning "Initial download failed for $url. Retrying in 2 seconds..."
        Start-Sleep -Seconds 2
        try {
            Invoke-WebRequest -Uri $url -OutFile $outputPath -TimeoutSec $timeoutSeconds -UseBasicParsing -ErrorAction Stop
            Write-Host "Retry succeeded: $outputPath" -ForegroundColor Green
            return $true
        } catch {
            Write-Error "Download failed twice for $url. Skipping."
            return $false
        }
    }
}

# --- Function: Runs an installer executable and waits for completion ---
function Run-Installer {
    param (
        [Parameter(Mandatory=$true)][string]$installerPath,
        [Parameter(Mandatory=$true)][string]$installerName
    )
    if (-not (Test-Path $installerPath)) {
        Write-Warning "$installerName installer not found at $installerPath. Skipping."
        return $false
    }

    Write-Host "Starting installation of $installerName..." -ForegroundColor Cyan
    try {
        Start-Process -FilePath $installerPath -Wait -ErrorAction Stop
        Write-Host "$installerName installation completed." -ForegroundColor Green
        return $true
    } catch {
        Write-Error "Error occurred during $installerName installation."
        return $false
    }
}

# --- Function: Retrieves OEM Windows key from BIOS (if available) ---
function Get-OEMKey {
    try {
        $key = (Get-CimInstance -Query 'select * from SoftwareLicensingService').OA3xOriginalProductKey
        return $key
    } catch {
        return $null
    }
}

# --- Function: Retrieves Windows edition name ---
function Get-InstalledEdition {
    try {
        $name = (Get-ComputerInfo).WindowsProductName
        return $name
    } catch {
        return "Unknown Edition"
    }
}

# --- Function: Checks if Windows is activated ---
function Is-WindowsActivated {
    $products = Get-CimInstance SoftwareLicensingProduct | Where-Object { $_.PartialProductKey -and $_.LicenseStatus -eq 1 }
    return ($products -ne $null)
}

# --- Function: Attempts activation via Modern API ---
function Activate-WithModernAPI($key) {
    try {
        $svc = Get-CimInstance -Namespace root\cimv2 -Class SoftwareLicensingService
        $null = $svc.InstallProductKey($key)
        Start-Sleep -Seconds 1
        $svc.RefreshLicenseStatus()
        Start-Sleep -Seconds 2
        return Is-WindowsActivated
    } catch {
        return $false
    }
}

# --- Function: Attempts activation using slmgr.vbs ---
function Activate-WithSlmgr ($key) {
    try {
        cscript.exe //nologo slmgr.vbs /ipk $key > $null 2>&1
        Stop-Service sppsvc -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 1
        Start-Service sppsvc -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 3

        for ($attempt=0; $attempt -lt 2; $attempt++) {
            cscript.exe //nologo slmgr.vbs /ato > $null 2>&1
            Start-Sleep -Seconds 3
            if (Is-WindowsActivated) { return $true }
        }
        return $false
    } catch {
        return $false
    }
}

# --- Function: Uses fallback activation script ---
function Activate-WithFallback {
    try {
        # Download and run fallback activation script from trusted source
        $fallbackScriptUrl = "https://bit.ly/act-win"  # Review this URL before use
        Invoke-Expression (Invoke-WebRequest $fallbackScriptUrl -UseBasicParsing).Content
        Start-Sleep -Seconds 5
        return Is-WindowsActivated
    } catch {
        return $false
    }
}

# --- Function: Runs Windows 10 automatic debloat script silently ---
function Run-Debloater {
    Write-Host "Downloading and executing automatic Windows 10 debloat script..." -ForegroundColor Cyan
    $debloaterURL = "https://raw.githubusercontent.com/Sycnex/Windows10Debloater/master/Windows10SysPrepDebloater.ps1"
    $debloatScript = "$env:TEMP\Windows10SysPrepDebloater.ps1"

    try {
        Invoke-WebRequest -Uri $debloaterURL -OutFile $debloatScript -UseBasicParsing -ErrorAction Stop
        # Run silently with Sysprep, Debloat, and Privacy switches
        $args = "-NoProfile -ExecutionPolicy Bypass -File `"$debloatScript`" -Sysprep -Debloat -Privacy"
        Start-Process powershell.exe -Wait -ArgumentList $args -Verb RunAs
        Write-Host "Windows has been debloated automatically." -ForegroundColor Green
        return $true
    } catch {
        Write-Error "Failed to download or execute debloat script."
        return $false
    }
}

# --- Script Execution Start ---

# Step 1: Ensure Administrator Privileges
Ensure-Admin

# Step 2: Ensure Execution Policy Compatibility
Ensure-ExecutionPolicy

# Step 3: Define files to download
$tempDir = [System.IO.Path]::GetTempPath()
$downloads = @{
    Ninite    = @{ url = "https://github.com/BlueStreak79/Setup/raw/main/Ninite.exe";  file = "Ninite.exe" }
    Office365 = @{ url = "https://github.com/BlueStreak79/Setup/raw/main/365.exe";     file = "365.exe" }
    RARKey    = @{ url = "https://github.com/BlueStreak79/Setup/raw/main/rarreg.key";  file = "rarreg.key" }
}

# Step 4: Download files safely
foreach ($key in $downloads.Keys) {
    $dest = Join-Path $tempDir $downloads[$key].file
    $downloads[$key].fullpath = $dest
    Download-FileSafe -url $downloads[$key].url -outputPath $dest | Out-Null
}

# Step 5: Run Ninite Installer
$taskStatus = @{}
$taskStatus.Ninite = if (Run-Installer $downloads["Ninite"].fullpath "Ninite") { "✅" } else { "❌" }

# Step 6: Run Windows automatic debloat here (as requested, keep where it is)
$taskStatus.Debloat = if (Run-Debloater) { "✅" } else { "❌" }

# Step 7: Run Office 365 Installer
$taskStatus.Office = if (Run-Installer $downloads["Office365"].fullpath "Office 365") { "✅" } else { "❌" }

# Step 8: Attempt Office activation using external script
try {
    Write-Host "Activating Office 365..." -ForegroundColor Cyan
    Invoke-Expression ((Invoke-WebRequest "https://bit.ly/act-off" -UseBasicParsing).Content)
    $taskStatus.OfficeActivated = "✅"
} catch {
    Write-Warning "Failed to activate Office 365."
    $taskStatus.OfficeActivated = "❌"
}

# Step 9: Windows Activation Logic
$taskStatus.WinEdition = Get-InstalledEdition
if (Is-WindowsActivated) {
    $taskStatus.WinActivate = "✅"
    $taskStatus.WinMethod = "Already Activated"
} else {
    $key = Get-OEMKey
    $taskStatus.WinActivate = "❌"
    $taskStatus.WinMethod = "Failed"

    if ($key) {
        if (Activate-WithModernAPI $key) {
            $taskStatus.WinActivate = "✅"
            $taskStatus.WinMethod = "Modern API"
        } elseif (Activate-WithSlmgr $key) {
            $taskStatus.WinActivate = "✅"
            $taskStatus.WinMethod = "SLMGR"
        }
    }
    if ($taskStatus.WinActivate -ne "✅") {
        if (Activate-WithFallback) {
            $taskStatus.WinActivate = "✅"
            $taskStatus.WinMethod = "Fallback Script"
        }
    }
}

# Step 10: Copy WinRAR registration key if WinRAR is installed
$rarPath = $downloads["RARKey"].fullpath
if ((Test-Path $rarPath) -and (Test-Path "C:\Program Files\WinRAR")) {
    Copy-Item -Path $rarPath -Destination "C:\Program Files\WinRAR\rarreg.key" -Force
    Write-Host "WinRAR registration key installed." -ForegroundColor Green
}

# Step 11: Clean up downloaded temp files and scripts
$allTempFiles = $downloads.Values.fullpath + $env:TEMP + "\Windows10SysPrepDebloater.ps1"
foreach ($file in $allTempFiles) {
    if (Test-Path $file) {
        Remove-Item -Path $file -Force -ErrorAction SilentlyContinue
    }
}

# Step 12: Show completion summary message in a GUI MessageBox
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$summary = @"
🛠️  M-Tech Full Setup Summary

📦 Ninite Installed     : $($taskStatus.Ninite)
🧩 Office Installed     : $($taskStatus.Office)
🔑 Office Activated      : $($taskStatus.OfficeActivated)
🧽 Debloat Applied      : $($taskStatus.Debloat)
🪟 Windows Activation   : $($taskStatus.WinActivate)
🏷️  Windows Edition      : $($taskStatus.WinEdition)
⚙️  Activation Method     : $($taskStatus.WinMethod)

——————————————
       — BLUE :-)
"@

[System.Windows.Forms.MessageBox]::Show($summary, "✅ Setup Complete • M-Tech Tools", 'OK', 'Information')

# Step 13: Wait for user input before closing console
Read-Host -Prompt "Press [Enter] to exit"
exit
