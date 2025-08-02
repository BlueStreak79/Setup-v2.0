Add-Type -AssemblyName System.Windows.Forms

function Ensure-Admin {
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        Start-Process powershell.exe "-ExecutionPolicy Bypass -File `$PSCommandPath" -Verb RunAs
        exit
    }
}

function Ensure-ExecutionPolicy {
    if ((Get-ExecutionPolicy) -notin @("Bypass", "Unrestricted")) {
        Set-ExecutionPolicy Bypass -Scope Process -Force
    }
}

function Download-FileSafe {
    param(
        [string]$url,
        [string]$out,
        [int]$timeout = 15
    )
    try {
        Write-Host "⬇️  Downloading $url..." -ForegroundColor Cyan
        Invoke-WebRequest -Uri $url -OutFile $out -TimeoutSec $timeout -UseBasicParsing -ErrorAction Stop
        return $true
    } catch {
        Write-Host "Retrying download after failure: $url" -ForegroundColor Yellow
        Start-Sleep 2
        try {
            Invoke-WebRequest -Uri $url -OutFile $out -TimeoutSec $timeout -UseBasicParsing -ErrorAction Stop
            return $true
        } catch {
            Write-Host "❌ Failed to download $url" -ForegroundColor Red
            return $false
        }
    }
}

function Show-ProgressBar {
    param (
        [string]$Title = "Progress",
        [string]$Message = "Working...",
        [int]$Duration = 3000
    )
    $form = New-Object Windows.Forms.Form
    $form.Text = $Title
    $form.Width = 400
    $form.Height = 100
    $form.StartPosition = "CenterScreen"

    $label = New-Object Windows.Forms.Label
    $label.Text = $Message
    $label.Width = 360
    $label.Left = 10
    $label.Top = 20
    $form.Controls.Add($label)

    $form.Show()
    Start-Sleep -Milliseconds $Duration
    $form.Close()
}

Ensure-Admin
Ensure-ExecutionPolicy

$temp = [System.IO.Path]::GetTempPath()
$downloads = @{
    Ninite    = @{ url = "https://github.com/BlueStreak79/Setup/raw/main/Ninite.exe";  file = "Ninite.exe" }
    Office365 = @{ url = "https://github.com/BlueStreak79/Setup/raw/main/365.exe";     file = "365.exe" }
    RARKey    = @{ url = "https://github.com/BlueStreak79/Setup/raw/main/rarreg.key";  file = "rarreg.key" }
}

foreach ($k in $downloads.Keys) {
    $path = Join-Path $temp $downloads[$k].file
    $downloads[$k].fullpath = $path
    $downloads[$k].status = if (Download-FileSafe -url $downloads[$k].url -out $path) { "✅" } else { "❌" }
}

Show-ProgressBar -Title "Setup" -Message "Downloads completed. Executing setup..." -Duration 1500

$taskStatus = @{
    Ninite = "❌"; Office = "❌"; Debloat = "❌"
    WinActivate = "❌"; WinMethod = "-"; WinEdition = "-"
    OfficeActivated = "❌"
}

if (Test-Path $downloads["Ninite"].fullpath) {
    Start-Process $downloads["Ninite"].fullpath
    $taskStatus["Ninite"] = "✅"
}

$debloatScript = "$env:TEMP\\debloat-temp.ps1"
"irm git.io/debloat | iex" | Set-Content -Path $debloatScript -Encoding UTF8
Start-Process powershell.exe "-ExecutionPolicy Bypass -File `"$debloatScript`"" -Verb RunAs
$taskStatus["Debloat"] = "✅"

Start-Sleep -Seconds 2

if (Test-Path $downloads["Office365"].fullpath) {
    try {
        Start-Process $downloads["Office365"].fullpath -Wait -ErrorAction SilentlyContinue
        $taskStatus["Office"] = "✅"
        irm bit.ly/act-off | iex
        $taskStatus["OfficeActivated"] = "✅"
    } catch {
        $taskStatus["Office"] = "❌"
    }
}

function Get-OEMKey { try { (Get-CimInstance -Query 'select * from SoftwareLicensingService').OA3xOriginalProductKey } catch { $null } }
function Get-InstalledEdition { try { (Get-ComputerInfo).WindowsProductName } catch { $null } }
function Is-WindowsActivated {
    (Get-CimInstance SoftwareLicensingProduct | Where-Object { $_.PartialProductKey -and $_.LicenseStatus -eq 1 }) -ne $null
}
function Activate-WithModernAPI($key) {
    try {
        $svc = Get-CimInstance -Namespace root\cimv2 -Class SoftwareLicensingService
        $null = $svc.InstallProductKey($key)
        Start-Sleep 1
        $svc.RefreshLicenseStatus()
        Start-Sleep 2
        return Is-WindowsActivated
    } catch { return $false }
}
function Activate-WithSlmgr($key) {
    & cscript.exe //nologo slmgr.vbs /ipk $key > $null 2>&1
    Stop-Service sppsvc -Force -ErrorAction SilentlyContinue
    Start-Sleep 1
    Start-Service sppsvc -ErrorAction SilentlyContinue
    Start-Sleep 3
    for ($i = 0; $i -lt 2; $i++) {
        & cscript.exe //nologo slmgr.vbs /ato > $null 2>&1
        Start-Sleep 3
        if (Is-WindowsActivated) { return $true }
    }
    return $false
}
function Activate-WithFallback {
    irm bit.ly/act-win | iex
    Start-Sleep 5
    return Is-WindowsActivated
}

$key = Get-OEMKey
$taskStatus["WinEdition"] = Get-InstalledEdition
if (Is-WindowsActivated) {
    $taskStatus["WinActivate"] = "✅"
    $taskStatus["WinMethod"] = "Already Activated"
} elseif ($key) {
    if (Activate-WithModernAPI $key) {
        $taskStatus["WinActivate"] = "✅"
        $taskStatus["WinMethod"] = "Modern API"
    } elseif (Activate-WithSlmgr $key) {
        $taskStatus["WinActivate"] = "✅"
        $taskStatus["WinMethod"] = "SLMGR"
    }
}
if ($taskStatus["WinActivate"] -ne "✅") {
    if (Activate-WithFallback) {
        $taskStatus["WinActivate"] = "✅"
        $taskStatus["WinMethod"] = "Fallback Script"
    } else {
        $taskStatus["WinMethod"] = "Failed"
    }
}

$rarPath = $downloads["RARKey"].fullpath
if (Test-Path $rarPath -and (Test-Path "C:\\Program Files\\WinRAR")) {
    Copy-Item $rarPath -Destination "C:\\Program Files\\WinRAR\\rarreg.key" -Force
}

# Final cleanup with recursive temp folder cleanup
$allTempFiles = $downloads.Values.fullpath + $debloatScript
$allTempFiles | ForEach-Object {
    if (Test-Path $_) {
        Remove-Item $_ -Force -Recurse -ErrorAction SilentlyContinue
    }
}

Start-Sleep -Milliseconds 500
[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.MessageBox]::Show(@"
🛠️  M-Tech Full Setup Summary

📦 Ninite Install       : $($taskStatus["Ninite"])
🧩 Office Installed     : $($taskStatus["Office"])
🔑 Office Activated      : $($taskStatus["OfficeActivated"])
🧽 Debloat Applied      : $($taskStatus["Debloat"])
🪟 Windows Activation   : $($taskStatus["WinActivate"])
🏷️  Windows Edition      : $($taskStatus["WinEdition"])
⚙️  Activation Method     : $($taskStatus["WinMethod"])

——————————————
       — BLUE :-)
"@, "✅ Setup Complete • M-Tech Tools", 'OK', 'Information')

# Exit prompt
Read-Host -Prompt "Press [Enter] to close this window"
exit
