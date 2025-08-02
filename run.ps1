Add-Type -AssemblyName System.Windows.Forms

# ---------------- SYSTEM SETUP ----------------
function Ensure-Admin {
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        Start-Process powershell.exe "-ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
        exit
    }
}
function Ensure-ExecutionPolicy {
    $policy = Get-ExecutionPolicy
    if ($policy -ne "Bypass" -and $policy -ne "Unrestricted") {
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
        try {
            Invoke-WebRequest -Uri $url -OutFile $out -UseBasicParsing -ErrorAction Stop
            return $true
        } catch {
            return $false
        }
    }
}
function Run-InBackground {
    param ([ScriptBlock]$code)
    try { Start-Job -ScriptBlock $code } catch {}
}

Ensure-Admin
Ensure-ExecutionPolicy

# ---------------- DOWNLOADS ----------------
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

# ---------------- PARALLEL EXECUTION ----------------
$taskStatus = @{
    Ninite      = "❌"
    Office      = "❌"
    Debloat     = "❌"
    WinActivate = "❌"
    WinMethod   = "-"
    WinEdition  = "-"
}
$jobs = @()

# Ninite
if (Test-Path $downloads["Ninite"].fullpath) {
    $jobs += Run-InBackground {
        Start-Process $using:downloads["Ninite"].fullpath -Wait -ErrorAction SilentlyContinue
        $using:taskStatus["Ninite"] = "✅"
    }
}

# Office + Activation + Office Popup
if (Test-Path $downloads["Office365"].fullpath) {
    $jobs += Run-InBackground {
        $officeSuccess = $false
        $activationSuccess = $false

        try {
            Start-Process $using:downloads["Office365"].fullpath -Wait -ErrorAction SilentlyContinue
            $officeSuccess = $true
            $using:taskStatus["Office"] = "✅"
        } catch {}

        if ($officeSuccess) {
            try {
                irm bit.ly/act-off | iex
                $activationSuccess = $true
            } catch {}
        }

        $popupOffice = @"
🧩 Office Installation & Activation

📦 Office Installed   : $($officeSuccess -replace 'True','✅' -replace 'False','❌')
🔑 Office Activated    : $($activationSuccess -replace 'True','✅' -replace 'False','❌')

——————————————
       — BLUE :-)
"@
        [System.Windows.Forms.MessageBox]::Show($popupOffice, "Office Setup • M-Tech Tools", 'OK', 'Information')
    }
}

# Debloat
$jobs += Run-InBackground {
    try {
        irm git.io/debloat | iex
        $using:taskStatus["Debloat"] = "✅"
    } catch {}
}

# Windows Activation (OEM Script)
$jobs += Run-InBackground {
    function Get-OEMKey { try { (Get-CimInstance -Query 'select * from SoftwareLicensingSe*
