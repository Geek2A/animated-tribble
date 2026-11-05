<#
.SYNOPSIS
    Removes the Purcheo print agent service and its files.

.EXAMPLE
    .\Uninstall-PurcheoPrintAgent.ps1
    .\Uninstall-PurcheoPrintAgent.ps1 -KeepConfig   # keep config.json (server URL and token) for a reinstall

.NOTES
    Run from an elevated PowerShell prompt. Reinstall with Install-PurcheoPrintAgent.ps1, which reuses
    a kept config.json. Custom page sizes the agent set on a printer are not reverted; reset them in the
    printer's preferences if needed. The agent token stays valid on the server; delete it in Purcheo
    (Settings > Printers > Print Agents) if the agent is retired.
#>
#Requires -RunAsAdministrator
[CmdletBinding()]
param(
    [string]$InstallDir = (Join-Path $env:ProgramData 'PurcheoPrintAgent'),
    [string]$ServiceName = 'PurcheoPrintAgent',
    [switch]$KeepConfig
)

$ErrorActionPreference = 'Stop'

$svc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
if ($svc) {
    if ($svc.Status -ne 'Stopped') {
        Stop-Service -Name $ServiceName -Force
        $svc.WaitForStatus('Stopped', [TimeSpan]::FromSeconds(30))
    }
    & sc.exe delete $ServiceName | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "sc.exe delete failed ($LASTEXITCODE)" }
    # A service open in services.msc stays marked for deletion until that window is closed.
    for ($i = 0; $i -lt 10 -and (Get-Service -Name $ServiceName -ErrorAction SilentlyContinue); $i++) { Start-Sleep -Seconds 1 }
    if (Get-Service -Name $ServiceName -ErrorAction SilentlyContinue) {
        Write-Warning "Service '$ServiceName' is marked for deletion; close services.msc and it will disappear."
    } else {
        Write-Host "Removed service '$ServiceName'"
    }
} else {
    Write-Host "Service '$ServiceName' not found"
}

# An agent left running by a crashed service host would keep files in the folder locked.
Get-CimInstance Win32_Process -Filter "Name = 'powershell.exe'" |
    Where-Object { $_.ProcessId -ne $PID -and $_.CommandLine -and $_.CommandLine -like "*$InstallDir*" } |
    ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
Get-Process -Name 'SumatraPDF' -ErrorAction SilentlyContinue |
    Where-Object { $_.Path -and $_.Path -like "$InstallDir*" } |
    Stop-Process -Force -ErrorAction SilentlyContinue

if (Test-Path $InstallDir) {
    if ($KeepConfig) {
        Get-ChildItem -Path $InstallDir -Force |
            Where-Object { $_.Name -ne 'config.json' } |
            Remove-Item -Recurse -Force
        Write-Host "Removed files from $InstallDir; kept config.json"
    } else {
        Remove-Item -Path $InstallDir -Recurse -Force
        Write-Host "Removed $InstallDir"
    }
}

# Temp PDFs from the LocalSystem account; the agent normally deletes them after each job.
$systemTemp = Join-Path $env:WINDIR 'Temp\PurcheoPrintAgent'
if (Test-Path $systemTemp) { Remove-Item -Path $systemTemp -Recurse -Force -ErrorAction SilentlyContinue }
