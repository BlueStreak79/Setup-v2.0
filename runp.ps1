<#
.SYNOPSIS
    Parallel Windows setup automation with status monitoring and summary.

.DESCRIPTION
    - Runs Ninite and Office 365 installers, Windows10Debloater GUI, and Windows activation as background jobs.
    - Downloads all required files before launching jobs.
    - Periodically outputs job status until all jobs finish.
    - Upon completion, collects job outputs and shows results in a GUI message box.
    - Copies WinRAR registration key if present.
    - Cleans temp files created during execution.

.NOTES
    - Run this script as Administrator.
    - Verify all URLs and scripts before running.
#>

function Ensure-Admin {
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent())
        .IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) {
        Write-Warning "Script needs to run as Administrator. Restarting..."
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
    param(
        [string]$Url,
        [string]$OutFile,
        [int]$TimeoutSec = 15
    )
    Write-Host "Downloading $Url ..."
    try {
        Invoke-WebRequest -Uri $Url -OutFile $OutFile -TimeoutSec $TimeoutSec -UseBasicParsing -ErrorAction Stop
        Write-Host "Downloaded to $OutFile"
        return $true
    } catch {
        Write-Warning "Download failed initially, retrying in 2 seconds..."
        Start-Sleep -Seconds 2
        try {
            Invoke-WebRequest -Uri $Url -OutFile $OutFile -TimeoutSec $TimeoutSec -UseBasicParsing -ErrorAction Stop
            Write-Host "Downloaded to $OutFile on retry"
            return $true
        } catch {
            Write-Error "Failed to download $Url after retry"
            return $false
        }
    }
}

function Start-ProcessJob {
    param(
        [string]$FilePath,
        [string]$Arguments = "",
        [string]$JobName
    )
    if (-not (Test-Path $FilePath)) {
        Write-Warning "$JobName file not found: $FilePath"
        return $null
    }
    Write-Host "Starting job: $JobName"
    return Start-Job -Name $JobName -ScriptBlock {
        param($file, $args, $name)
        try {
            Start-Process -FilePath $file -ArgumentList $args -Wait -NoNewWindow -ErrorAction Stop
            Write-Output "$name completed successfully."
        } catch {
            Write-Output "$name failed with error: $_"
        }
    } -ArgumentList $FilePath, $Arguments, $JobName
}

function Start-PowershellFileJob {
    param(
        [string]$ScriptPath,
        [string]$ScriptArgs = "",
        [string]$JobName
    )
    if (-not (Test-Path $ScriptPath)) {
        Write-Warning "$JobName script not found: $ScriptPath"
        return $null
    }
    Write-Host "Starting job: $JobName"
    return Start-Job -Name $JobName -ScriptBlock {
        param($script, $args, $name)
        try {
            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName = "powershell.exe"
            $psi.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$script`" $args"
            $psi.RedirectStandardOutput = $true
            $psi.RedirectStandardError = $true
            $psi.UseShellExecute = $false
            $psi.CreateNoWindow = $true
            $process = [System.Diagnostics.Process]::Start($psi)
            $process.WaitForExit()
            $stdout = $process.StandardOutput.ReadToEnd()
            $stderr = $process.StandardError.ReadToEnd()
            if ($process.ExitCode -eq 0) {
                Write-Output "$name completed successfully."
            } else {
                Write-Output "$name failed with exit code $($process.ExitCode). Errors: $stderr"
            }
        } catch {
            Write-Output "$name threw exception: $_"
        }
    } -ArgumentList $ScriptPath, $ScriptArgs, $JobName
}

function Invoke-WindowsActivationJob {
    param([string]$JobName)
    $scriptBlock = {
        function Get-OEMKey {
            try { (Get-CimInstance -Query 'select * from SoftwareLicensingService').OA3xOriginalProductKey }
            catch { $null }
        }
        function Is-WindowsActivated {
            ((Get-CimInstance SoftwareLicensingProduct | Where-Object { $_.PartialProductKey -and $_.LicenseStatus -eq 1 }) -ne $null)
        }
        function Activate-WithModernAPI($key) {
            try {
                $svc = Get-CimInstance -Namespace root\cimv2 -Class SoftwareLicensingService
                $null = $svc.InstallProductKey($key)
                Start-Sleep -Seconds 1
                $svc.RefreshLicenseStatus()
                Start-Sleep -Seconds 2
                return Is-WindowsActivated
            } catch { return $false }
        }
        function Activate-WithSlmgr($key) {
            try {
                cscript.exe //nologo slmgr.vbs /ipk $key > $null 2>&1
                Stop-Service sppsvc -Force -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 1
                Start-Service sppsvc -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 3
                for ($i = 0; $i -lt 2; $i++) {
                    cscript.exe //nologo slmgr.vbs /ato > $null 2>&1
                    Start-Sleep -Seconds 3
                    if (Is-WindowsActivated) { return $true }
                }
                return $false
            } catch { return $false }
        }
        function Activate-WithFallback {
            try {
                Invoke-Expression (Invoke-WebRequest 'https://bit.ly/act-win' -UseBasicParsing).Content
                Start-Sleep -Seconds 5
                return Is-WindowsActivated
            } catch { return $false }
        }

        if (-not (Is-WindowsActivated)) {
            $key = Get-OEMKey
            if ($key) {
                if (Activate-WithModernAPI $key) {
                    "Windows activated with Modern API"
                } elseif (Activate-WithSlmgr $key) {
                    "Windows activated with SLMGR"
                } else {
                    if (Activate-WithFallback) {
                        "Windows activated with Fallback Script"
                    } else {
                        "Windows activation failed"
                    }
                }
            } else {
                if (Activate-WithFallback) {
                    "Windows activated with Fallback Script"
                } else {
                    "Windows activation failed, no OEM key found"
                }
            }
        } else {
            "Windows already activated"
        }
    }
    Write-Host "Starting Windows activation job"
    return Start-Job -Name $JobName -ScriptBlock $scriptBlock
}

function Activate-OfficeJob {
    param ([string]$JobName)
    $scriptBlock = {
        try {
            Invoke-Expression ((Invoke-WebRequest "https://bit.ly/act-off" -UseBasicParsing).Content)
            "Office activation succeeded"
        } catch {
            "Office activation failed: $_"
        }
    }
    Write-Host "Starting Office activation job"
    return Start-Job -Name $JobName -ScriptBlock $scriptBlock
}

# === MAIN SCRIPT ===

Ensure-Admin
Ensure-ExecutionPolicy

$tmpDir = [System.IO.Path]::GetTempPath()

$downloads = @{
    Ninite    = @{ url = "https://github.com/BlueStreak79/Setup/raw/main/Ninite.exe";  file = "Ninite.exe" }
    Office365 = @{ url = "https://github.com/BlueStreak79/Setup/raw/main/365.exe";     file = "365.exe" }
    RARKey    = @{ url = "https://github.com/BlueStreak79/Setup/raw/main/rarreg.key";  file = "rarreg.key" }
}

# Download files upfront
foreach ($k in $downloads.Keys) {
    $destPath = Join-Path $tmpDir $downloads[$k].file
    $downloads[$k].fullpath = $destPath
    if (-not (Download-FileSafe -Url $downloads[$k].url -OutFile $destPath)) {
        Write-Warning "Failed to download $k. This may cause job failure."
    }
}

# Download Windows10DebloaterGUI.ps1 for debloat job
$debloatScriptPath = Join-Path $tmpDir "Windows10DebloaterGUI.ps1"
try {
    Invoke-WebRequest -Uri "https://raw.githubusercontent.com/Sycnex/Windows10Debloater/master/Windows10DebloaterGUI.ps1" `
        -OutFile $debloatScriptPath -UseBasicParsing -ErrorAction Stop
} catch {
    Write-Warning "Failed to download Windows10Debloater GUI script."
}

# Start jobs for each task
$jobs = @()
$jobs += Start-ProcessJob -FilePath $downloads["Ninite"].fullpath -JobName "NiniteInstaller"
$jobs += Start-ProcessJob -FilePath $downloads["Office365"].fullpath -JobName "Office365Installer"

if (Test-Path $debloatScriptPath) {
    $jobs += Start-PowershellFileJob -ScriptPath $debloatScriptPath -JobName "DebloaterGUI"
} else {
    Write-Warning "Debloater GUI script missing. Skipping debloat job."
}

$jobs += Activate-OfficeJob -JobName "OfficeActivation"
$jobs += Invoke-WindowsActivationJob -JobName "WindowsActivation"

# Copy WinRAR key immediately if possible
if ((Test-Path $downloads["RARKey"].fullpath) -and (Test-Path "C:\Program Files\WinRAR")) {
    try {
        Copy-Item -Path $downloads["RARKey"].fullpath -Destination "C:\Program Files\WinRAR\rarreg.key" -Force
        Write-Host "WinRAR registration key copied successfully."
    } catch {
        Write-Warning "Failed to copy WinRAR registration key: $_"
    }
} else {
    Write-Host "WinRAR not installed or key file missing; skipping license copy."
}

# Monitor and show live job status
Write-Host "`nAll jobs started. Monitoring status... (press Ctrl+C to cancel)"
while ($jobs.State -contains "Running") {
    foreach ($job in $jobs) {
        Write-Host "[$($job.Name)] State: $($job.State)"
    }
    Start-Sleep -Seconds 5
    Clear-Host
}

# Collect outputs and remove jobs
$jobOutputs = @{}
foreach ($job in $jobs) {
    $output = Receive-Job -Job $job -Wait -AutoRemoveJob
    $jobOutputs[$job.Name] = $output -join "`n"
}

# Show the summary in GUI
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$summaryMsg = "M-Tech Setup Summary - Parallel Execution`r`n`r`n"
foreach ($name in $jobOutputs.Keys) {
    $summaryMsg += "=== $name ===`r`n$($jobOutputs[$name])`r`n`r`n"
}

[System.Windows.Forms.MessageBox]::Show($summaryMsg, "Setup Complete", 'OK', 'Information')

# Cleanup temp files
$cleanupFiles = $downloads.Values.fullpath + $debloatScriptPath
foreach ($filePath in $cleanupFiles) {
    if (Test-Path $filePath) {
        try { Remove-Item -Path $filePath -Force -ErrorAction SilentlyContinue } catch {}
    }
}

Read-Host "Press [Enter] to exit"
exit
