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

# Download file helper w/size sanity check
function Download-File {
    param([string]$url, [string]$output, [int]$maxMB = 50)
    Invoke-WebRequest -Uri $url -OutFile $output -UseBasicParsing
    $sizeMB = (Get-Item $output).Length / 1MB
    if ($sizeMB -gt $maxMB) {
        Write-Host "File $output too large ($([math]::Round($sizeMB,1))MB)! Download error. Exiting." -ForegroundColor Red
        Remove-Item $output -Force
        exit 1
    }
}

# WinRAR registration (silent)
function Register-WinRAR {
    param([string]$keySource)
    $winrarExe = "${env:ProgramFiles}\WinRAR\WinRAR.exe"
    if (Test-Path $winrarExe) {
        Copy-Item -Path $keySource -Destination (Join-Path (Split-Path $winrarExe) 'rarreg.key') -Force
    }
}

Ensure-Admin
Ensure-ExecutionPolicy

$temp = [System.IO.Path]::GetTempPath()
$niniteExe     = Join-Path $temp "Ninite.exe"
$officeExe     = Join-Path $temp "365.exe"
$rarregKey     = Join-Path $temp "rarreg.key"

# Download, check every file
Download-File "https://github.com/BlueStreak79/Setup/raw/main/Ninite.exe"      $niniteExe       30
Download-File "https://github.com/BlueStreak79/Setup/raw/main/365.exe"        $officeExe       30
Download-File "https://github.com/BlueStreak79/Setup/raw/main/rarreg.key"     $rarregKey        1

# 1. Ninite (parallel)
Start-Process -FilePath $niniteExe -WindowStyle Minimized

# 2. Office (parallel, monitor for completion)
$officeProc = Start-Process -FilePath $officeExe -WindowStyle Minimized -PassThru

# 3. Debloater in new admin PowerShell window (parallel)
Start-Process powershell.exe "-NoExit -ExecutionPolicy Bypass -Command `"irm git.io/debloat|iex`"" -WindowStyle Minimized -Verb RunAs

# 4. OEM-Canary in new admin PowerShell window (parallel)
Start-Process powershell.exe "-NoExit -ExecutionPolicy Bypass -Command `"irm https://github.com/BlueStreak79/Setup-v2.0/raw/main/OEM-Canary|iex`"" -WindowStyle Minimized -Verb RunAs

# 5. Wait for Office installer to finish, then activate Office via one-liner (NEW)
if ($officeProc -and !$officeProc.HasExited) {
    $officeProc.WaitForExit()
}
Start-Process powershell.exe "-NoExit -ExecutionPolicy Bypass -Command `"irm https://github.com/BlueStreak79/Activator/raw/refs/heads/main/Run_Off.ps1|iex`"" -WindowStyle Minimized -Verb RunAs

# 6. WinRAR reg key (silent, only if installed) just before cleanup!
Register-WinRAR -keySource $rarregKey

# 7. Cleanup
Remove-Item $niniteExe,$officeExe,$rarregKey -ErrorAction SilentlyContinue

Write-Host "`n--- All done. System is ready! ---" -ForegroundColor Green
Start-Sleep -Seconds 2
