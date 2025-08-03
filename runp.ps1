# Ensure admin elevation
function Ensure-Admin {
    if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Start-Process powershell.exe "-File `"$PSCommandPath`"" -Verb RunAs
        exit
    }
}

# Set script execution policy for this process
function Ensure-ExecutionPolicy {
    $policy = Get-ExecutionPolicy
    if ($policy -ne 'Unrestricted') {
        Set-ExecutionPolicy Unrestricted -Scope Process -Force
    }
}

# Download file helper
function Download-File {
    param ([string]$url, [string]$output)
    Invoke-WebRequest -Uri $url -OutFile $output
}

# Replace WinRAR reg key if installed (silent)
function Register-WinRAR {
    param ([string]$keySource)
    $winrarExe = "${env:ProgramFiles}\WinRAR\WinRAR.exe"
    if (Test-Path $winrarExe) {
        Copy-Item -Path $keySource -Destination (Join-Path (Split-Path $winrarExe) 'rarreg.key') -Force
    }
}

# MAIN SCRIPT
Ensure-Admin
Ensure-ExecutionPolicy

$temp = [System.IO.Path]::GetTempPath()
$niniteExe = Join-Path $temp "Ninite.exe"
$officeExe = Join-Path $temp "365.exe"
$rarregKey = Join-Path $temp "rarreg.key"
$oemCanaryExe = Join-Path $temp "OEM-Canary.exe"

# Download all needed files
Download-File "https://github.com/BlueStreak79/Setup/raw/main/Ninite.exe" $niniteExe
Download-File "https://github.com/BlueStreak79/Setup/raw/main/365.exe"    $officeExe
Download-File "https://github.com/BlueStreak79/Setup/raw/main/rarreg.key" $rarregKey
Download-File "https://github.com/BlueStreak79/Setup-v2.0/raw/refs/heads/main/OEM-Canary" $oemCanaryExe

# 1. Launch Ninite (parallel, minimized)
Start-Process -FilePath $niniteExe -WindowStyle Minimized

# 2. Launch Office installer (parallel, minimized). We'll monitor for when it closes for activation.
$officeProc = Start-Process -FilePath $officeExe -WindowStyle Minimized -PassThru

# 3. Launch debloat script in a new window, parallel
Start-Process powershell.exe "-NoExit -Command `"iwr -useb git.io/debloat | iex`"" -WindowStyle Minimized

# 4. Launch OEM-Canary in a new window, parallel
Start-Process -FilePath $oemCanaryExe -WindowStyle Minimized

# 5. WinRAR registration (if present)
Register-WinRAR -keySource $rarregKey

# 6. Wait for Office installer, then activate
if ($officeProc -and !$officeProc.HasExited) {
    $officeProc.WaitForExit()
}
Start-Process powershell.exe "-ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -File `"$temp\ActivateOffice.ps1`"" -NoNewWindow

# Download office activation script just in time
$activateScript = Join-Path $temp "ActivateOffice.ps1"
Download-File "https://github.com/BlueStreak79/Activator/raw/refs/heads/main/Run_Win.ps1" $activateScript

# Run activation (already triggered above after Office installed)

# 7. Cleanup
$downloads = @($niniteExe, $officeExe, $rarregKey, $oemCanaryExe, $activateScript)
foreach ($f in $downloads) {
    if (Test-Path $f) { Remove-Item $f -Force }
}

# 8. Finish / Visual indicator
Write-Host "`n--- All done. System is ready! ---" -ForegroundColor Green
Start-Sleep -Seconds 2
