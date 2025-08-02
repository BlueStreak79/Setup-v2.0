Add-Type -AssemblyName System.Windows.Forms

# Ensure Admin & Execution Policy
function Ensure-Admin {
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        Start-Process powershell.exe "-ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
        exit
    }
}
function Ensure-ExecutionPolicy {
    if ((Get-ExecutionPolicy) -notin @("Bypass", "Unrestricted")) {
        Set-ExecutionPolicy Bypass -Scope Process -Force
    }
}
function Download-FileSafe {
    param($url, $out)
    try {
        Invoke-WebRequest -Uri $url -OutFile $out -UseBasicParsing -ErrorAction Stop
        return $true
    } catch {
        Start-Sleep 2
        try { Invoke-WebRequest -Uri $url -OutFile $out -UseBasicParsing -ErrorAction Stop } catch { return $false }
    }
}
function Run-InBackground { param([ScriptBlock]$code); Start-Job -ScriptBlock $code }

Ensure-Admin
Ensure-ExecutionPolicy

# Downloads
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

# Track Task Status
$taskStatus = @{
    Ninite = "❌"; Office = "❌"; Debloat = "❌"
    WinActivate = "❌"; WinMethod = "-"; WinEdition = "-"
    OfficeActivated = "❌"
}

# Ninite Installer
$niniteJob = $null
if (Test-Path $downloads["Ninite"].fullpath) {
    $niniteJob = Run-InBackground {
        Start-Process $using:downloads["Ninite"].fullpath -Wait -ErrorAction SilentlyContinue
    }
}

# Debloat
$debloatJob = Run-InBackground {
    try { irm git.io/debloat | iex } catch {}
}

# Office + Office Activation + Popup
$officeJob = $null
if (Test-Path $downloads["Office365"].fullpath) {
    $officeJob = Run-InBackground {
        $result = @{ Installed = $false; Activated = $false }
        try {
            Start-Process $using:downloads["Office365"].fullpath -Wait -ErrorAction SilentlyContinue
            $result.Installed = $true
            irm bit.ly/act-off | iex
            $result.Activated = $true
        } catch {}
        $result
    }
}

# ---------------- Run Windows Activation inline (safe for vars) ----------------
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

# Run OEM Activation Now
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

# ---------------- Wait for Jobs & Gather Results ----------------
if ($niniteJob) { $niniteJob | Wait-Job | Receive-Job; $taskStatus["Ninite"] = "✅" }
$debloatJob | Wait-Job | Receive-Job; $taskStatus["Debloat"] = "✅"

if ($officeJob) {
    $result = $officeJob | Wait-Job | Receive-Job
    $taskStatus["Office"] = if ($result.Installed) { "✅" } else { "❌" }
    $taskStatus["OfficeActivated"] = if ($result.Activated) { "✅" } else { "❌" }

    # Office Popup
    $popupOffice = @"
🧩 Office Installation & Activation

📦 Office Installed   : $($taskStatus["Office"])
🔑 Office Activated    : $($taskStatus["OfficeActivated"])

——————————————
       — BLUE :-)
"@
    [System.Windows.Forms.MessageBox]::Show($popupOffice, "Office Setup • M-Tech Tools", 'OK', 'Information')
}

# ---------------- Final Summary ----------------
$popup = @"
🛠️  M-Tech Setup Summary

📦 Ninite Install     : $($taskStatus["Ninite"])
🧩 Office Install     : $($taskStatus["Office"])
🧽 Debloat Applied    : $($taskStatus["Debloat"])
🪟 Windows Activation : $($taskStatus["WinActivate"])
🏷️  Win Edition       : $($taskStatus["WinEdition"])
🔧 Method Used        : $($taskStatus["WinMethod"])

——————————————
       — BLUE :-)
"@
[System.Windows.Forms.MessageBox]::Show($popup, "Setup Completed • M-Tech Tools", 'OK', 'Information')

# Optional: Clean up downloaded files
$downloads.Values | ForEach-Object {
    $f = $_.fullpath
    if (Test-Path $f) { Remove-Item $f -Force -ErrorAction SilentlyContinue }
}
