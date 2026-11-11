#Requires -Version 5.1
<#
.SYNOPSIS
    Blue Toolkit - single-file Windows setup, OEM licensing, diagnostics and network repair.

.DESCRIPTION
    Works two ways:
      1) Locally:      .\BlueToolkit.ps1                 (menu)
                       .\BlueToolkit.ps1 -List           (task keys)
                       .\BlueToolkit.ps1 -Task full-setup
      2) From GitHub:  irm <raw-url-of-this-file> | iex  (set scriptUrl below so it can self-elevate)
                       $env:TOOLKIT_TASK = 'full-setup'; irm <raw-url> | iex

    Edit the CONFIG block, add your own steps in the USER EXTENSIONS region, or drop a custom.ps1
    next to this file (local runs only). Run it on a test VM before trusting it on a real machine.
#>

#region CONFIG ============================================================================
$script:Config = @{
    scriptUrl = ''     # raw URL of THIS file; only needed to self-elevate when launched via irm | iex

    assets = @{
        ninite = @{ url = 'https://github.com/BlueStreak79/Setup/raw/main/Ninite.exe'; sha256 = ''; maxMB = 30 }
        office = @{ url = 'https://github.com/BlueStreak79/Setup/raw/main/365.exe';    sha256 = ''; maxMB = 30 }
    }

    debloat = @{
        disableConsumerFeatures = $true
        apps = @(
            'Microsoft.BingNews', 'Microsoft.BingWeather', 'Microsoft.GetHelp', 'Microsoft.Getstarted',
            'Microsoft.MicrosoftSolitaireCollection', 'Microsoft.WindowsFeedbackHub', 'Microsoft.People',
            'Microsoft.MicrosoftOfficeHub', 'Clipchamp.Clipchamp'
        )
        url = ''; sha256 = ''     # optional extra debloat script you have reviewed (pin a hash)
    }

    winrar = @{ keyPath = '' }    # path to a WinRAR key file you supply; nothing is bundled
    winget = @{ packages = @() }  # e.g. @('7zip.7zip', 'VideoLAN.VLC')

    settings = @{
        timezone = 'India Standard Time'
        lightTheme = $true
        showDesktopIcons = $true
        taskbarSearchIconOnly = $true
        explorerOpensThisPC = $true
        hideQuickAccessHistory = $true
    }

    hosts = @{ tag = 'toolkit'; domains = @() }

    diagnostics = @{
        tools = @(
            @{ key = '1'; name = 'AquaKeyTest';     url = 'https://github.com/BlueStreak79/Diagnostics/raw/main/AquaKeyTest.exe';     sha256 = '' },
            @{ key = '2'; name = 'BatteryInfoView'; url = 'https://github.com/BlueStreak79/Diagnostics/raw/main/BatteryInfoView.exe'; sha256 = '' },
            @{ key = '3'; name = 'Camera';          url = 'https://github.com/BlueStreak79/Diagnostics/raw/main/Camera.exe';          sha256 = '' },
            @{ key = '4'; name = 'lusrmgr';         url = 'https://github.com/BlueStreak79/Diagnostics/raw/main/lusrmgr.exe';         sha256 = '' },
            @{ key = '5'; name = 'MemTest64';       url = 'https://github.com/BlueStreak79/Diagnostics/raw/main/MemTest64.exe';       sha256 = '' },
            @{ key = '6'; name = 'LCDtest';         url = 'https://github.com/BlueStreak79/Diagnostics/raw/main/LCDtest.exe';         sha256 = '' },
            @{ key = '7'; name = 'OemKey';          url = 'https://github.com/BlueStreak79/Diagnostics/raw/main/OemKey.exe';          sha256 = '' },
            @{ key = '8'; name = 'ShowKeyPlus';     url = 'https://github.com/BlueStreak79/Diagnostics/raw/main/ShowKeyPlus.exe';     sha256 = '' }
        )
    }
}
function Get-ToolkitConfig { $script:Config }
#endregion

#region CORE ==============================================================================
# Common.ps1 - shared helpers. ASCII only so Windows PowerShell 5.1 reads it correctly.

$script:LogDir  = Join-Path $env:TEMP 'Toolkit\logs'
$script:WorkDir = Join-Path $env:TEMP 'Toolkit\work'
$script:Tasks   = [ordered]@{}
$script:Results = New-Object System.Collections.ArrayList

foreach ($d in @($script:LogDir, $script:WorkDir)) {
    if (-not (Test-Path $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
}
$script:LogFile = Join-Path $script:LogDir ("toolkit_{0}.log" -f (Get-Date -Format 'yyyyMMdd'))

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','OK','WARN','ERROR')][string]$Level = 'INFO'
    )
    $color = @{ INFO = 'Cyan'; OK = 'Green'; WARN = 'Yellow'; ERROR = 'Red' }[$Level]
    $line  = "[{0}] [{1}] {2}" -f (Get-Date -Format 'HH:mm:ss'), $Level, $Message
    Write-Host $line -ForegroundColor $color
    try { Add-Content -Path $script:LogFile -Value $line -ErrorAction Stop } catch { }
}

function Test-IsAdmin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    ([Security.Principal.WindowsPrincipal]$id).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Save-VerifiedDownload {
    <# Download with retry, size cap and optional SHA256 pin. Returns $true/$false. #>
    param(
        [Parameter(Mandatory)][string]$Url,
        [Parameter(Mandatory)][string]$OutFile,
        [string]$Sha256,
        [int]$MaxMB = 50,
        [int]$Retries = 2
    )
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    $ProgressPreference = 'SilentlyContinue'
    for ($i = 0; $i -le $Retries; $i++) {
        try {
            Write-Log "Downloading $Url"
            Invoke-WebRequest -Uri $Url -OutFile $OutFile -UseBasicParsing -TimeoutSec 120 -ErrorAction Stop
            $mb = (Get-Item $OutFile).Length / 1MB
            if ($mb -gt $MaxMB) {
                Remove-Item $OutFile -Force
                throw ("file is {0:N1} MB, over the {1} MB limit" -f $mb, $MaxMB)
            }
            if ($Sha256) {
                $h = (Get-FileHash -Path $OutFile -Algorithm SHA256).Hash
                if ($h -ne $Sha256.ToUpper()) {
                    Remove-Item $OutFile -Force
                    throw "SHA256 mismatch (got $h)"
                }
            } else {
                Write-Log ("No SHA256 pinned for {0} - UNVERIFIED download" -f (Split-Path $OutFile -Leaf)) 'WARN'
            }
            return $true
        } catch {
            Write-Log ("Download attempt {0} failed: {1}" -f ($i + 1), $_.Exception.Message) 'WARN'
            Start-Sleep -Seconds 2
        }
    }
    return $false
}

function Add-StepResult {
    param(
        [Parameter(Mandatory)][string]$Step,
        [Parameter(Mandatory)][ValidateSet('OK','Skipped','Failed')][string]$Status,
        [string]$Message = ''
    )
    [void]$script:Results.Add([pscustomobject]@{ Step = $Step; Status = $Status; Message = $Message })
    $lvl = switch ($Status) { 'OK' { 'OK' } 'Skipped' { 'INFO' } default { 'ERROR' } }
    Write-Log ("{0}: {1} {2}" -f $Step, $Status, $Message) $lvl
}

function Show-Summary {
    param([switch]$Popup)
    Write-Host ''
    $script:Results | Format-Table -AutoSize | Out-Host
    if ($Popup -and $script:Results.Count -gt 0) {
        Add-Type -AssemblyName System.Windows.Forms
        $text = ($script:Results | ForEach-Object { "{0,-26} {1,-8} {2}" -f $_.Step, $_.Status, $_.Message }) -join "`r`n"
        [void][System.Windows.Forms.MessageBox]::Show($text, 'Toolkit summary', 'OK', 'Information')
    }
}

function Register-ToolkitTask {
    <# Modules and custom\*.ps1 files call this to appear in the menu / -Task. #>
    param(
        [Parameter(Mandatory)][string]$Key,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][scriptblock]$Action
    )
    $script:Tasks[$Key] = [pscustomobject]@{ Key = $Key; Name = $Name; Action = $Action }
}

function Invoke-ToolkitTask {
    param([Parameter(Mandatory)][string]$Key)
    $t = $script:Tasks.Values | Where-Object { $_.Key -ieq $Key } | Select-Object -First 1
    if (-not $t) { Write-Log "Unknown task '$Key'. Use -List to see tasks." 'ERROR'; return }
    Write-Log "Running task: $($t.Name)"
    try { & $t.Action }
    catch { Write-Log "Task '$($t.Key)' crashed: $($_.Exception.Message)" 'ERROR'; Add-StepResult $t.Name 'Failed' $_.Exception.Message }
}

$script:Hooks = @{ PreSetup = @(); PostSetup = @() }

function Register-ToolkitHook {
    <# Run your own scriptblock at the start (PreSetup) or end (PostSetup) of the full-setup task. #>
    param(
        [Parameter(Mandatory)][ValidateSet('PreSetup', 'PostSetup')][string]$Stage,
        [Parameter(Mandatory)][scriptblock]$Action
    )
    $script:Hooks[$Stage] += $Action
}

function Invoke-ToolkitHooks {
    param([Parameter(Mandatory)][string]$Stage)
    foreach ($h in @($script:Hooks[$Stage])) {
        if (-not $h) { continue }
        try { & $h } catch { Add-StepResult "Hook ($Stage)" 'Failed' $_.Exception.Message }
    }
}

function Invoke-SelfElevate {
    <# Relaunches elevated. Works for a saved file (-File) or for irm | iex when scriptUrl is set. #>
    param([string[]]$ArgList = @())
    if ($PSCommandPath) {
        $a = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $PSCommandPath)) + $ArgList
    }
    elseif ($script:Config.scriptUrl) {
        $cmd = "irm '{0}' | iex" -f $script:Config.scriptUrl
        if ($ArgList.Count -ge 2 -and $ArgList[0] -eq '-Task') { $cmd = ("`$env:TOOLKIT_TASK='{0}'; " -f $ArgList[1]) + $cmd }
        $a = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', ('"{0}"' -f $cmd))
    }
    else {
        Write-Host 'Not running as administrator. Open PowerShell with "Run as administrator" and start it again,' -ForegroundColor Yellow
        Write-Host '(or set scriptUrl in the CONFIG block so it can elevate itself).' -ForegroundColor Yellow
        return $false
    }
    Start-Process -FilePath powershell.exe -ArgumentList $a -Verb RunAs
    return $true
}
#endregion

#region MODULE 00-FullSetup ==========================================================
# 00-FullSetup.ps1 - the "do everything" run. Long jobs start together; dependent steps wait only on what they need.

function Invoke-FullSetup {
    [CmdletBinding()] param()
    $cfg = Get-ToolkitConfig
    Invoke-ToolkitHooks -Stage 'PreSetup'

    # 1. Quick, independent steps
    Install-WingetIfMissing
    Set-WindowsPreferences
    Remove-BloatApps

    # 2. Start the slow things at the same time (none of these block)
    $ninite = Start-AssetInstaller -Key 'ninite' -Label 'Ninite'
    $office = Start-AssetInstaller -Key 'office' -Label 'Office installer'
    $debloat = Start-DebloatScript

    # 3. OEM activation only needs internet, so it can run while installers work
    $r = Invoke-OemActivation
    if ($r.Success)             { Add-StepResult 'OEM activation' 'OK' $r.Method }
    elseif (-not $r.KeyFound)   { Add-StepResult 'OEM activation' 'Skipped' $r.Message }
    else                        { Add-StepResult 'OEM activation' 'Failed' $r.Message }

    # 4. WinRAR key must wait for Ninite (which installs WinRAR)
    Wait-ToolkitHandles -Handles @($ninite)
    Register-WinRARKey

    # 5. Everything else
    Wait-ToolkitHandles -Handles @($office, $debloat)

    # 6. winget packages after the installers finish (avoids installer-mutex clashes)
    Install-WingetPackages -Ids @($cfg.winget.packages)

    Invoke-ToolkitHooks -Stage 'PostSetup'
    Clear-ToolkitWork
    Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
    Start-Process explorer.exe
    Show-Summary -Popup
}

Register-ToolkitTask -Key 'full-setup' -Name 'FULL SETUP (winget, settings, apps, OEM activation, cleanup)' -Action { Invoke-FullSetup }
#endregion

#region MODULE 10-Winget =============================================================
# 10-Winget.ps1 - WinGet bootstrap (from Winget repo) + package installs.

function Install-WingetIfMissing {
    [CmdletBinding()] param()
    $ProgressPreference = 'SilentlyContinue'

    if (Get-Command winget -ErrorAction SilentlyContinue) {
        Add-StepResult 'WinGet bootstrap' 'Skipped' 'Already installed'
        return
    }

    $admin = Test-IsAdmin
    $scope = if ($admin) { 'AllUsers' } else { 'CurrentUser' }
    try {
        Write-Log 'Ensuring NuGet package provider...'
        Install-PackageProvider -Name NuGet -Force -Confirm:$false | Out-Null

        Write-Log "Installing Microsoft.WinGet.Client module ($scope)..."
        Install-Module -Name Microsoft.WinGet.Client -Repository PSGallery -Force -Confirm:$false -AllowClobber -Scope $scope | Out-Null

        if ($admin) {
            Write-Log 'Bootstrapping WinGet (all users)...'
            Repair-WinGetPackageManager -AllUsers
        } else {
            Write-Log 'Not admin: skipping Repair-WinGetPackageManager.' 'WARN'
        }

        if (Get-Command winget -ErrorAction SilentlyContinue) {
            Add-StepResult 'WinGet bootstrap' 'OK' 'Installed'
        } else {
            Add-StepResult 'WinGet bootstrap' 'Failed' 'Not detected afterwards (LTSC / Store-less system?)'
        }
    } catch {
        Add-StepResult 'WinGet bootstrap' 'Failed' $_.Exception.Message
    }
}

function Install-WingetPackages {
    [CmdletBinding()]
    param([string[]]$Ids)

    if (-not $Ids -or $Ids.Count -eq 0) {
        Add-StepResult 'WinGet packages' 'Skipped' 'No packages listed in config.json (winget.packages)'
        return
    }
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        Add-StepResult 'WinGet packages' 'Failed' 'winget not available'
        return
    }

    $failed = @()
    foreach ($id in $Ids) {
        Write-Log "winget install $id"
        & winget install --id $id --exact --silent --accept-package-agreements --accept-source-agreements
        # -1978335189 = already installed / no applicable update
        if ($LASTEXITCODE -ne 0 -and $LASTEXITCODE -ne -1978335189) { $failed += $id }
    }
    if ($failed.Count -eq 0) { Add-StepResult 'WinGet packages' 'OK' ("{0} package(s)" -f $Ids.Count) }
    else { Add-StepResult 'WinGet packages' 'Failed' ("failed: " + ($failed -join ', ')) }
}

Register-ToolkitTask -Key 'winget-bootstrap' -Name 'WinGet: install / repair' -Action { Install-WingetIfMissing }
Register-ToolkitTask -Key 'winget-packages'  -Name 'WinGet: install packages from config.json' -Action {
    Install-WingetPackages -Ids @((Get-ToolkitConfig).winget.packages)
}
#endregion

#region MODULE 20-Apps ===============================================================
# 20-Apps.ps1 - Ninite, Office installer, optional debloat script, WinRAR key from a local file.
# Installers are started without blocking so they run in parallel; Wait-ToolkitHandles waits for them.

function Start-AssetInstaller {
    param(
        [Parameter(Mandatory)][string]$Key,     # key under "assets" in config.json
        [Parameter(Mandatory)][string]$Label
    )
    $cfg = (Get-ToolkitConfig).assets.$Key
    if (-not $cfg -or -not $cfg.url) { Add-StepResult $Label 'Skipped' 'No URL in config.json'; return $null }

    $out = Join-Path $script:WorkDir (Split-Path $cfg.url -Leaf)
    $max = if ($cfg.maxMB) { [int]$cfg.maxMB } else { 50 }
    if (-not (Save-VerifiedDownload -Url $cfg.url -OutFile $out -Sha256 $cfg.sha256 -MaxMB $max)) {
        Add-StepResult $Label 'Failed' 'Download failed'
        return $null
    }
    $p = Start-Process -FilePath $out -PassThru
    [pscustomobject]@{ Label = $Label; Process = $p; File = $out }
}

function Start-DebloatScript {
    <# Downloads the script named in config.json (debloat.url) and runs it hidden in its own PowerShell.
       Left blank by default: pin a real URL + SHA256 you have reviewed instead of a short link. #>
    $cfg = (Get-ToolkitConfig).debloat
    if (-not $cfg -or -not $cfg.url) { Add-StepResult 'Debloat' 'Skipped' 'No URL in config.json (debloat.url)'; return $null }

    $out = Join-Path $script:WorkDir 'debloat.ps1'
    if (-not (Save-VerifiedDownload -Url $cfg.url -OutFile $out -Sha256 $cfg.sha256 -MaxMB 5)) {
        Add-StepResult 'Debloat' 'Failed' 'Download failed'
        return $null
    }
    $p = Start-Process powershell.exe -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-WindowStyle', 'Hidden', '-File', ('"{0}"' -f $out)) -PassThru
    [pscustomobject]@{ Label = 'Debloat'; Process = $p; File = $out }
}

function Wait-ToolkitHandles {
    param([object[]]$Handles)
    foreach ($h in @($Handles)) {
        if (-not $h) { continue }
        Write-Log "Waiting for $($h.Label)..."
        $h.Process.WaitForExit()
        if ($h.Process.ExitCode -eq 0) { Add-StepResult $h.Label 'OK' }
        else { Add-StepResult $h.Label 'Failed' ("exit code {0}" -f $h.Process.ExitCode) }
    }
}

function Register-WinRARKey {
    <# Copies a key file YOU supply (config.json: winrar.keyPath) into the WinRAR folder. Nothing is bundled. #>
    $path = (Get-ToolkitConfig).winrar.keyPath
    if (-not $path) { Add-StepResult 'WinRAR key' 'Skipped' 'winrar.keyPath is empty'; return }
    if (-not (Test-Path $path)) { Add-StepResult 'WinRAR key' 'Failed' "key file not found: $path"; return }

    $dir = Join-Path $env:ProgramFiles 'WinRAR'
    if (-not (Test-Path (Join-Path $dir 'WinRAR.exe'))) { Add-StepResult 'WinRAR key' 'Skipped' 'WinRAR is not installed'; return }

    Copy-Item -Path $path -Destination (Join-Path $dir 'rarreg.key') -Force
    Add-StepResult 'WinRAR key' 'OK'
}

function Clear-ToolkitWork {
    Get-ChildItem -Path $script:WorkDir -File -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
}

Register-ToolkitTask -Key 'apps' -Name 'Apps: Ninite + Office installer (parallel), then WinRAR key' -Action {
    $h = @((Start-AssetInstaller -Key 'ninite' -Label 'Ninite'), (Start-AssetInstaller -Key 'office' -Label 'Office installer')) | Where-Object { $_ }
    Wait-ToolkitHandles -Handles $h
    Register-WinRARKey
    Clear-ToolkitWork
}
#endregion

#region MODULE 30-Settings ===========================================================
# 30-Settings.ps1 - Windows preferences (from WIndows-Settings repo), driven by config.json "settings".
# Note: HKCU settings apply to the account the script runs as.

function Set-RegValue {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)]$Value,
        [string]$Type = 'DWord'
    )
    if (-not (Test-Path $Path)) { New-Item -Path $Path -Force | Out-Null }
    New-ItemProperty -Path $Path -Name $Name -Value $Value -PropertyType $Type -Force | Out-Null
}

function Set-WindowsPreferences {
    [CmdletBinding()]
    param([switch]$RestartExplorer)

    $s = (Get-ToolkitConfig).settings
    $cv = 'HKCU:\Software\Microsoft\Windows\CurrentVersion'
    try {
        if ($s.timezone) {
            & tzutil.exe /s "$($s.timezone)"
            Write-Log "Time zone set to $($s.timezone)"
        }
        if ($s.lightTheme) {
            Set-RegValue "$cv\Themes\Personalize" 'SystemUsesLightTheme' 1
            Set-RegValue "$cv\Themes\Personalize" 'AppsUseLightTheme' 1
            Write-Log 'Light theme applied'
        }
        if ($s.showDesktopIcons) {
            $p = "$cv\Explorer\HideDesktopIcons\NewStartPanel"
            foreach ($g in @(
                '{20D04FE0-3AEA-1069-A2D8-08002B30309D}',   # This PC
                '{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}',   # Network
                '{59031a47-3f72-44a7-89c5-5595fe6b30ee}',   # User files
                '{645FF040-5081-101B-9F08-00AA002F954E}'    # Recycle Bin
            )) { Set-RegValue $p $g 0 }
            Write-Log 'Desktop icons enabled'
        }
        if ($s.taskbarSearchIconOnly) {
            Set-RegValue "$cv\Search" 'SearchboxTaskbarMode' 1
            Write-Log 'Taskbar search set to icon only'
        }
        if ($s.explorerOpensThisPC) {
            Set-RegValue "$cv\Explorer\Advanced" 'LaunchTo' 1
            Write-Log 'Explorer opens to This PC'
        }
        if ($s.hideQuickAccessHistory) {
            Set-RegValue "$cv\Explorer\Advanced" 'ShowFrequent' 0
            Set-RegValue "$cv\Explorer\Advanced" 'ShowRecent' 0
            Write-Log 'Quick Access history hidden'
        }
        if ($RestartExplorer) {
            Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
            Start-Process explorer.exe
        }
        Add-StepResult 'Windows preferences' 'OK'
    } catch {
        Add-StepResult 'Windows preferences' 'Failed' $_.Exception.Message
    }
}

Register-ToolkitTask -Key 'settings' -Name 'Windows preferences (theme, icons, Explorer, time zone)' -Action {
    Set-WindowsPreferences -RestartExplorer
}

function Remove-BloatApps {
    <# Conservative built-in debloat: disables auto-installed promo apps and removes the apps listed in
       config "debloat.apps". Edit that list to taste. #>
    [CmdletBinding()] param()
    $d = (Get-ToolkitConfig).debloat
    try {
        if ($d.disableConsumerFeatures) {
            Set-RegValue 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent' 'DisableWindowsConsumerFeatures' 1
            Write-Log 'Consumer features (auto-installed promo apps) disabled'
        }
        $removed = 0
        foreach ($name in @($d.apps)) {
            if (-not $name) { continue }
            try {
                $pk = Get-AppxPackage -Name $name -AllUsers -ErrorAction SilentlyContinue
                if ($pk) { $pk | Remove-AppxPackage -AllUsers -ErrorAction Stop; $removed++ }
                Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue |
                    Where-Object { $_.DisplayName -eq $name } |
                    Remove-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue | Out-Null
            } catch {
                Write-Log ("Could not remove {0}: {1}" -f $name, $_.Exception.Message) 'WARN'
            }
        }
        Add-StepResult 'Debloat (apps)' 'OK' ("{0} app(s) removed" -f $removed)
    } catch {
        Add-StepResult 'Debloat (apps)' 'Failed' $_.Exception.Message
    }
}

Register-ToolkitTask -Key 'debloat' -Name 'Debloat: remove listed apps + disable promo auto-installs' -Action { Remove-BloatApps }
#endregion

#region MODULE 40-OemLicense =========================================================
<#
.SYNOPSIS
    Reads the OEM (OA3/MSDM) Windows key embedded in firmware and activates with it.

.DESCRIPTION
    Works only on machines whose BIOS/UEFI carries a Windows key (most branded
    laptops/desktops sold with Windows). VMs, retail and custom-built PCs
    usually have none, in which case the module reports that and changes nothing.

    Flow:
      1. Read the firmware key (SoftwareLicensingService.OA3xOriginalProductKey)
      2. Check whether Windows is already activated
      3. Install the key, refresh licensing, activate (CIM first, slmgr /ato as fallback)
      4. Return a result object so a caller can build a summary

.EXAMPLE
    .\OemLicense.ps1 -ReportOnly          # show key + license state, change nothing
    .\OemLicense.ps1                      # activate with the firmware key
    .\OemLicense.ps1 -ShowKey             # include the full key in the output

.NOTES
    Dot-source to use as a module function:  . .\OemLicense.ps1; Invoke-OemActivation
#>
$script:WindowsAppId = '55c92734-d682-4d71-983e-d6ec3f16059f'

function Test-IsAdmin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    ([Security.Principal.WindowsPrincipal]$id).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-OemProductKey {
    <# Returns the firmware-embedded key, or $null if the machine has none. #>
    try {
        $key = (Get-CimInstance -ClassName SoftwareLicensingService -ErrorAction Stop).OA3xOriginalProductKey
    } catch {
        return $null
    }
    if ([string]::IsNullOrWhiteSpace($key)) { return $null }
    $key.Trim()
}

function Format-MaskedKey {
    param([string]$Key)
    if ([string]::IsNullOrEmpty($Key) -or $Key.Length -lt 5) { return $null }
    'XXXXX-XXXXX-XXXXX-XXXXX-' + $Key.Substring($Key.Length - 5)
}

function Get-WindowsLicenseProduct {
    <# Raw CIM instance of the Windows license currently in use (has a partial key). #>
    Get-CimInstance -ClassName SoftwareLicensingProduct `
        -Filter "ApplicationId='$script:WindowsAppId' AND PartialProductKey IS NOT NULL" `
        -ErrorAction SilentlyContinue | Select-Object -First 1
}

function Get-WindowsLicenseInfo {
    $p = Get-WindowsLicenseProduct
    $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction SilentlyContinue
    [pscustomobject]@{
        Edition      = if ($os) { $os.Caption } else { $null }
        LicenseName  = if ($p) { $p.Name } else { $null }
        Channel      = if ($p) { $p.ProductKeyChannel } else { $null }   # OEM:DM, Retail, Volume:GVLK ...
        PartialKey   = if ($p) { $p.PartialProductKey } else { $null }
        LicenseState = if ($p) { $p.LicenseStatus } else { $null }       # 1 = Licensed
        IsActivated  = [bool]($p -and $p.LicenseStatus -eq 1)
    }
}

function Invoke-OemActivation {
    [CmdletBinding()]
    param(
        [switch]$ReportOnly,
        [switch]$ShowKey,
        [switch]$Force
    )

    $r = [ordered]@{
        KeyFound  = $false
        Key       = $null
        Edition   = $null
        Channel   = $null
        WasActive = $false
        IsActive  = $false
        Method    = 'None'
        Success   = $false
        Message   = ''
    }

    $info = Get-WindowsLicenseInfo
    $r.Edition   = $info.Edition
    $r.Channel   = $info.Channel
    $r.WasActive = $info.IsActivated
    $r.IsActive  = $info.IsActivated

    $key = Get-OemProductKey
    if (-not $key) {
        $r.Message = 'No OEM key found in firmware (VM, retail, or custom-built machine).'
        $r.Success = $info.IsActivated
        return [pscustomobject]$r
    }

    $r.KeyFound = $true
    $r.Key = if ($ShowKey) { $key } else { Format-MaskedKey $key }

    if ($ReportOnly) {
        $r.Method  = 'ReportOnly'
        $r.Success = $true
        $r.Message = 'Report only - nothing changed.'
        return [pscustomobject]$r
    }

    if ($info.IsActivated -and -not $Force) {
        $r.Method  = 'AlreadyActivated'
        $r.Success = $true
        $r.Message = 'Windows is already activated. Use -Force to re-apply the OEM key.'
        return [pscustomobject]$r
    }

    if (-not (Test-IsAdmin)) {
        $r.Message = 'Administrator rights are required to install a key.'
        return [pscustomobject]$r
    }

    # --- Install the key via the licensing service ---
    $svc = Get-CimInstance -ClassName SoftwareLicensingService
    try {
        $null = Invoke-CimMethod -InputObject $svc -MethodName InstallProductKey `
                    -Arguments @{ ProductKey = $key } -ErrorAction Stop
        $null = Invoke-CimMethod -InputObject $svc -MethodName RefreshLicenseStatus -ErrorAction SilentlyContinue
    } catch {
        $r.Message = "Key was rejected: $($_.Exception.Message). " +
                     'Most often this is an edition mismatch (e.g. a Home key on a Pro install).'
        return [pscustomobject]$r
    }

    # --- Activate: CIM first ---
    try {
        $prod = Get-WindowsLicenseProduct
        if ($prod) {
            $null = Invoke-CimMethod -InputObject $prod -MethodName Activate -ErrorAction Stop
        }
    } catch { }
    Start-Sleep -Seconds 2
    if ((Get-WindowsLicenseInfo).IsActivated) {
        $r.Method = 'CIM'
    }
    else {
        # --- Fallback: slmgr /ato ---
        $slmgr = Join-Path $env:windir 'System32\slmgr.vbs'
        & cscript.exe //nologo $slmgr /ato 2>&1 | Out-Null
        Start-Sleep -Seconds 3
        if ((Get-WindowsLicenseInfo).IsActivated) { $r.Method = 'slmgr' }
    }

    $final = Get-WindowsLicenseInfo
    $r.IsActive = $final.IsActivated
    $r.Channel  = $final.Channel
    $r.Success  = $final.IsActivated
    $r.Message  = if ($final.IsActivated) {
        'Activated with the firmware OEM key.'
    } else {
        'Key installed but activation did not complete. Check internet access and that the key matches the installed edition.'
    }
    [pscustomobject]$r
}

Register-ToolkitTask -Key 'oem-report' -Name 'OEM key: show firmware key + license state (no changes)' -Action {
    Invoke-OemActivation -ReportOnly -ShowKey | Format-List | Out-Host
}

Register-ToolkitTask -Key 'oem-activate' -Name 'OEM key: activate Windows with the firmware key' -Action {
    $r = Invoke-OemActivation
    $r | Format-List | Out-Host
    if ($r.Success)        { Add-StepResult 'OEM activation' 'OK' $r.Method }
    elseif (-not $r.KeyFound) { Add-StepResult 'OEM activation' 'Skipped' $r.Message }
    else                   { Add-StepResult 'OEM activation' 'Failed' $r.Message }
}
#endregion

#region MODULE 50-Diagnostics ========================================================
# 50-Diagnostics.ps1 - diagnostics menu (Diagnostics repo + Test/Cam-Audio.ps1), driven by config.json.

function Show-SystemInfo {
    [CmdletBinding()] param([switch]$Popup)

    $cs    = Get-CimInstance Win32_ComputerSystem
    $os    = Get-CimInstance Win32_OperatingSystem
    $cpu   = Get-CimInstance Win32_Processor | Select-Object -First 1
    $bios  = Get-CimInstance Win32_BIOS
    $board = Get-CimInstance Win32_BaseBoard
    $gpu   = (Get-CimInstance Win32_VideoController | Select-Object -ExpandProperty Name) -join ', '
    $disks = (Get-PhysicalDisk | ForEach-Object { "{0} ({1} GB)" -f $_.FriendlyName, [math]::Round($_.Size / 1GB) }) -join ', '

    $lines = @(
        'SYSTEM INFORMATION',
        '------------------------------',
        "Computer name : $env:COMPUTERNAME",
        "Model         : $($cs.Manufacturer) $($cs.Model)",
        "Serial number : $($bios.SerialNumber)",
        "Motherboard   : $($board.Manufacturer) $($board.Product)",
        "CPU           : $($cpu.Name)",
        ("RAM           : {0} GB" -f [math]::Round($cs.TotalPhysicalMemory / 1GB)),
        "Storage       : $disks",
        "GPU           : $gpu",
        "OS            : $($os.Caption) ($($os.Version))",
        "Architecture  : $env:PROCESSOR_ARCHITECTURE"
    )
    $lines | ForEach-Object { Write-Host $_ }

    if ($Popup) {
        Add-Type -AssemblyName System.Windows.Forms
        [void][System.Windows.Forms.MessageBox]::Show(($lines -join "`r`n"), 'System Information', 'OK', 'Information')
    }
}

function Invoke-DiagnosticTool {
    param([Parameter(Mandatory)]$Tool)
    $out = Join-Path $script:WorkDir ("{0}.exe" -f $Tool.name)
    if (-not (Save-VerifiedDownload -Url $Tool.url -OutFile $out -Sha256 $Tool.sha256 -MaxMB 60)) {
        Write-Log "Could not download $($Tool.name)" 'ERROR'
        return
    }
    Write-Log "Launching $($Tool.name)"
    Start-Process -FilePath $out
}

function Open-CameraAndSound {
    try { Start-Process 'microsoft.windows.camera:' -ErrorAction Stop } catch { Write-Log "Camera app failed: $($_.Exception.Message)" 'ERROR' }
    try { Start-Process control.exe -ArgumentList '/name Microsoft.Sound' -ErrorAction Stop } catch { Write-Log "Sound settings failed: $($_.Exception.Message)" 'ERROR' }
}

function Show-DiagnosticsMenu {
    $tools = @((Get-ToolkitConfig).diagnostics.tools)
    while ($true) {
        Clear-Host
        Write-Host '=== Diagnostics ===' -ForegroundColor Cyan
        foreach ($t in $tools) { Write-Host ("[{0}] {1}" -f $t.key, $t.name) }
        Write-Host '[S] System information'
        Write-Host '[C] Camera + Sound settings'
        Write-Host '[0] Back'
        $c = Read-Host 'Select'
        if ($c -eq '0') { return }
        elseif ($c -ieq 'S') { Show-SystemInfo -Popup }
        elseif ($c -ieq 'C') { Open-CameraAndSound }
        else {
            $tool = $tools | Where-Object { $_.key -eq $c } | Select-Object -First 1
            if ($tool) { Invoke-DiagnosticTool -Tool $tool } else { Write-Log 'Invalid choice' 'WARN' }
        }
        Read-Host 'Press ENTER to continue' | Out-Null
    }
}

Register-ToolkitTask -Key 'diagnostics' -Name 'Diagnostics tools menu' -Action { Show-DiagnosticsMenu }
Register-ToolkitTask -Key 'sysinfo'     -Name 'System information'     -Action { Show-SystemInfo -Popup }
#endregion

#region MODULE 60-Network ============================================================
# 60-Network.ps1 - Network Doctor (ported from Network-Diagnostics/Net.ps1).
# Changes from the original: no hard-coded placeholder URL for re-elevation (Run.ps1 handles that),
# destructive actions ask first, and the TCP step no longer disables auto-tuning or enables the
# long-deprecated "chimney" offload.

function Confirm-Action {
    param([Parameter(Mandatory)][string]$Prompt)
    (Read-Host "$Prompt (Y/N)") -match '^[Yy]$'
}

function Get-ActiveAdapter {
    Get-NetAdapter | Where-Object { $_.Status -eq 'Up' } | Select-Object -First 1
}

function Set-DnsServers {
    param([Parameter(Mandatory)][string]$Label, [string[]]$Servers)
    $a = Get-ActiveAdapter
    if (-not $a) { Write-Log 'No active network adapter found.' 'ERROR'; return }
    if ($Servers) { Set-DnsClientServerAddress -InterfaceAlias $a.Name -ServerAddresses $Servers }
    else          { Set-DnsClientServerAddress -InterfaceAlias $a.Name -ResetServerAddresses }
    Write-Log "$Label applied on $($a.Name)" 'OK'
}

function Reset-Winsock  { Write-Log 'Resetting Winsock...';        netsh winsock reset }
function Reset-IPStack  { Write-Log 'Resetting TCP/IP stack...';   netsh int ip reset; netsh interface ipv4 reset; netsh interface ipv6 reset }
function Clear-DnsCache { Write-Log 'Flushing DNS...';             ipconfig /flushdns }
function Renew-DhcpLease { Write-Log 'Renewing DHCP lease...';     ipconfig /release; ipconfig /renew }
function Reset-WinHttpProxy { Write-Log 'Resetting proxy...';      netsh winhttp reset proxy }
function Reset-WinFirewall  { Write-Log 'Resetting Windows Firewall...'; netsh advfirewall reset }
function Clear-ArpCache { Write-Log 'Clearing ARP cache...';       netsh interface ip delete arpcache }
function Reset-RouteTable { Write-Log 'Clearing routing table...'; route -f }
function Restart-AllAdapters { Write-Log 'Restarting adapters...'; Get-NetAdapter | Restart-NetAdapter -Confirm:$false }

function Invoke-FullNetworkRepair {
    Write-Host 'This runs: Winsock reset, TCP/IP reset, DNS flush, DHCP renew, proxy reset,' -ForegroundColor Yellow
    Write-Host 'firewall reset (custom rules are lost), ARP + route clear, adapter restart.' -ForegroundColor Yellow
    if (-not (Confirm-Action 'Continue?')) { return }
    Reset-Winsock; Reset-IPStack; Clear-DnsCache; Renew-DhcpLease; Reset-WinHttpProxy
    Reset-WinFirewall; Clear-ArpCache; Reset-RouteTable; Restart-AllAdapters
    Write-Log 'Full repair completed. A reboot is recommended.' 'OK'
}

function Invoke-FullStackReset {
    Write-Host 'WARNING: removes and reinstalls ALL network adapters (netcfg -d).' -ForegroundColor Red
    if (Confirm-Action 'Continue?') { netcfg -d; Write-Log 'Network stack rebuilt. Reboot recommended.' 'OK' }
}

function Remove-AllWifiProfiles {
    Write-Host 'WARNING: removes ALL saved Wi-Fi profiles.' -ForegroundColor Red
    if (Confirm-Action 'Continue?') { netsh wlan delete profile name=* }
}

function Remove-OneWifiProfile {
    $n = Read-Host 'Wi-Fi profile name'
    if ($n) { netsh wlan delete profile name="$n" }
}

function Set-TcpDefaults {
    Write-Log 'Applying TCP settings (auto-tuning normal, RSS on, CTCP)...'
    netsh int tcp set global autotuninglevel=normal
    netsh int tcp set global rss=enabled
    netsh int tcp set supplemental template=internet congestionprovider=ctcp
}

function Show-PublicIp {
    try { Write-Host ("Public IP: {0}" -f (Invoke-WebRequest -Uri 'https://api.ipify.org' -UseBasicParsing).Content) -ForegroundColor Green }
    catch { Write-Log 'Unable to fetch public IP' 'ERROR' }
}

# ---- Ookla Speedtest CLI ----
$script:SpeedDir = Join-Path $script:WorkDir 'Speedtest'
$script:SpeedExe = Join-Path $script:SpeedDir 'speedtest.exe'

function Install-SpeedtestCli {
    if (Test-Path $script:SpeedExe) { return $true }
    New-Item -ItemType Directory -Path $script:SpeedDir -Force | Out-Null
    $zip = Join-Path $script:WorkDir 'speedtest.zip'
    if (-not (Save-VerifiedDownload -Url 'https://install.speedtest.net/app/cli/ookla-speedtest-1.2.0-win64.zip' -OutFile $zip -MaxMB 20)) { return $false }
    Expand-Archive -Path $zip -DestinationPath $script:SpeedDir -Force
    Remove-Item $zip -Force
    Test-Path $script:SpeedExe
}

function Invoke-SpeedTest {
    param([switch]$Json)
    if (-not (Install-SpeedtestCli)) { Write-Log 'Speedtest CLI unavailable' 'ERROR'; return }
    if ($Json) {
        $p = Join-Path $script:LogDir ("SpeedTest_{0}.json" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
        & $script:SpeedExe --accept-license --accept-gdpr --format=json | Out-File -FilePath $p -Encoding ascii
        Write-Log "JSON report saved: $p" 'OK'
    } else {
        & $script:SpeedExe --accept-license --accept-gdpr
    }
}

$script:NetMenu = [ordered]@{
    '1'  = @{ Name = 'Full network repair';        Action = { Invoke-FullNetworkRepair } }
    '2'  = @{ Name = 'Winsock reset';              Action = { Reset-Winsock } }
    '3'  = @{ Name = 'TCP/IP reset';               Action = { Reset-IPStack } }
    '4'  = @{ Name = 'Flush DNS';                  Action = { Clear-DnsCache } }
    '5'  = @{ Name = 'Renew DHCP';                 Action = { Renew-DhcpLease } }
    '6'  = @{ Name = 'Firewall reset';             Action = { if (Confirm-Action 'Reset firewall rules?') { Reset-WinFirewall } } }
    '7'  = @{ Name = 'Proxy reset';                Action = { Reset-WinHttpProxy } }
    '8'  = @{ Name = 'Restart adapters';           Action = { Restart-AllAdapters } }
    '9'  = @{ Name = 'Full network stack reset';   Action = { Invoke-FullStackReset } }
    '10' = @{ Name = 'Google DNS';                 Action = { Set-DnsServers 'Google DNS' @('8.8.8.8', '8.8.4.4') } }
    '11' = @{ Name = 'Cloudflare DNS';             Action = { Set-DnsServers 'Cloudflare DNS' @('1.1.1.1', '1.0.0.1') } }
    '12' = @{ Name = 'Quad9 DNS';                  Action = { Set-DnsServers 'Quad9 DNS' @('9.9.9.9', '149.112.112.112') } }
    '13' = @{ Name = 'Restore DHCP DNS';           Action = { Set-DnsServers 'DHCP DNS' $null } }
    '14' = @{ Name = 'Show Wi-Fi profiles';        Action = { netsh wlan show profiles } }
    '15' = @{ Name = 'Delete all Wi-Fi profiles';  Action = { Remove-AllWifiProfiles } }
    '16' = @{ Name = 'Delete one Wi-Fi profile';   Action = { Remove-OneWifiProfile } }
    '17' = @{ Name = 'Run SFC';                    Action = { sfc /scannow } }
    '18' = @{ Name = 'Run DISM RestoreHealth';     Action = { DISM /Online /Cleanup-Image /RestoreHealth } }
    '19' = @{ Name = 'TCP settings';               Action = { Set-TcpDefaults } }
    '20' = @{ Name = 'Connectivity test';          Action = { Test-NetConnection google.com } }
    '21' = @{ Name = 'Ping';                       Action = { $h = Read-Host 'Host'; if ($h) { ping $h } } }
    '22' = @{ Name = 'Trace route';                Action = { $h = Read-Host 'Host'; if ($h) { tracert $h } } }
    '23' = @{ Name = 'DNS lookup';                 Action = { $h = Read-Host 'Domain'; if ($h) { nslookup $h } } }
    '24' = @{ Name = 'Public IP';                  Action = { Show-PublicIp } }
    '25' = @{ Name = 'Adapter information';        Action = { Get-NetAdapter | Format-Table -AutoSize | Out-Host } }
    '26' = @{ Name = 'Ookla speed test';           Action = { Invoke-SpeedTest } }
    '27' = @{ Name = 'Speed test (JSON export)';   Action = { Invoke-SpeedTest -Json } }
}

function Show-NetworkMenu {
    while ($true) {
        Clear-Host
        Write-Host "=== Network Doctor ===" -ForegroundColor Cyan
        foreach ($k in $script:NetMenu.Keys) { Write-Host ("{0,3}  {1}" -f $k, $script:NetMenu[$k].Name) }
        Write-Host '  0  Back'
        $c = Read-Host 'Select'
        if ($c -eq '0') { return }
        if ($script:NetMenu.Contains($c)) { & $script:NetMenu[$c].Action } else { Write-Log 'Invalid selection' 'WARN' }
        Read-Host 'Press ENTER to continue' | Out-Null
    }
}

Register-ToolkitTask -Key 'network' -Name 'Network Doctor menu' -Action { Show-NetworkMenu }
#endregion

#region MODULE 70-HostsFile ==========================================================
# 70-HostsFile.ps1 - generic hosts-file block manager. Domains come from config.json (hosts.domains).
# It writes a tagged block between markers, so re-running replaces the block instead of duplicating lines,
# and removing it leaves the rest of the file untouched. A backup is made before every change.

function Set-HostsBlock {
    [CmdletBinding()]
    param(
        [string[]]$Domains = @(),
        [string]$Tag = 'toolkit',
        [string]$Address = '0.0.0.0'
    )
    $hosts = Join-Path $env:SystemRoot 'System32\drivers\etc\hosts'
    Copy-Item -Path $hosts -Destination "$hosts.toolkit.bak" -Force

    $begin = "# >>> toolkit:$Tag"
    $end   = "# <<< toolkit:$Tag"
    $out   = New-Object System.Collections.ArrayList
    $skip  = $false
    foreach ($line in @(Get-Content -Path $hosts)) {
        if ($line -eq $begin) { $skip = $true;  continue }
        if ($line -eq $end)   { $skip = $false; continue }
        if (-not $skip) { [void]$out.Add($line) }
    }
    if ($Domains.Count -gt 0) {
        [void]$out.Add($begin)
        foreach ($d in ($Domains | Sort-Object -Unique)) { [void]$out.Add("$Address $d") }
        [void]$out.Add($end)
    }
    Set-Content -Path $hosts -Value $out -Encoding ASCII
    ipconfig /flushdns | Out-Null
}

Register-ToolkitTask -Key 'hosts-apply' -Name 'Hosts file: apply block list from config.json' -Action {
    $c = (Get-ToolkitConfig).hosts
    $domains = @($c.domains)
    if ($domains.Count -eq 0) { Add-StepResult 'Hosts block' 'Skipped' 'hosts.domains is empty in config.json'; return }
    Set-HostsBlock -Domains $domains -Tag $c.tag
    Add-StepResult 'Hosts block' 'OK' ("{0} domain(s)" -f $domains.Count)
}

Register-ToolkitTask -Key 'hosts-remove' -Name 'Hosts file: remove the toolkit block' -Action {
    Set-HostsBlock -Domains @() -Tag (Get-ToolkitConfig).hosts.tag
    Add-StepResult 'Hosts block removed' 'OK'
}
#endregion

#region USER EXTENSIONS ===================================================================
# Put your own steps here (this region is yours; nothing below it needs to change).
#
# Add a menu item / -Task key:
#   Register-ToolkitTask -Key 'my-task' -Name 'What it does' -Action {
#       # your code
#       Add-StepResult 'My task' 'OK' 'optional note'
#   }
#
# Run something automatically inside full-setup:
#   Register-ToolkitHook -Stage PostSetup -Action {
#       # runs after apps/settings, before cleanup and the summary
#   }
#
# Helpers available: Write-Log, Test-IsAdmin, Save-VerifiedDownload (retry + size cap + SHA256 pin),
#                    Get-ToolkitConfig, Add-StepResult, Show-Summary.
#endregion

#region MAIN ==============================================================================
# Arguments are parsed by hand (no param block) so the same file also works through irm | iex.
$TaskName  = $env:TOOLKIT_TASK
$ListOnly  = $false
$NoElevate = $false
for ($i = 0; $i -lt $args.Count; $i++) {
    switch -Regex ($args[$i]) {
        '^-Task$'      { if ($i + 1 -lt $args.Count) { $TaskName = $args[$i + 1]; $i++ } }
        '^-List$'      { $ListOnly = $true }
        '^-NoElevate$' { $NoElevate = $true }
    }
}

# Optional external file next to the script (local runs only)
if ($PSScriptRoot) {
    $customFile = Join-Path $PSScriptRoot 'custom.ps1'
    if (Test-Path $customFile) {
        try { . $customFile } catch { Write-Log ("custom.ps1 failed to load: {0}" -f $_.Exception.Message) 'ERROR' }
    }
}

function Show-MainMenu {
    while ($true) {
        Clear-Host
        Write-Host '==================================' -ForegroundColor Cyan
        Write-Host '          BLUE TOOLKIT            ' -ForegroundColor Cyan
        Write-Host '==================================' -ForegroundColor Cyan
        $n = 0
        $map = @{}
        foreach ($t in $script:Tasks.Values) {
            $n++
            $map["$n"] = $t
            Write-Host ("{0,3}  {1}" -f $n, $t.Name)
        }
        Write-Host '  0  Exit'
        $c = Read-Host 'Select'
        if ($c -eq '0') { return }
        if ($map.ContainsKey($c)) {
            Invoke-ToolkitTask -Key $map[$c].Key
            Show-Summary
        } else {
            Write-Log 'Invalid selection' 'WARN'
        }
        Read-Host 'Press ENTER to return to the menu' | Out-Null
    }
}

if ($ListOnly) {
    $script:Tasks.Values | ForEach-Object { "{0,-18} {1}" -f $_.Key, $_.Name }
}
elseif (-not $NoElevate -and -not (Test-IsAdmin)) {
    $extra = @()
    if ($TaskName) { $extra += @('-Task', $TaskName) }
    [void](Invoke-SelfElevate -ArgList $extra)
}
elseif ($TaskName) {
    Invoke-ToolkitTask -Key $TaskName
    Show-Summary
    Read-Host 'Press ENTER to close' | Out-Null
}
else {
    Show-MainMenu
}
#endregion
