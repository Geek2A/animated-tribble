<#
.SYNOPSIS
    Installs the Purcheo print agent as a Windows service. Safe to re-run to reinstall or upgrade.

.EXAMPLE
    .\Install-PurcheoPrintAgent.ps1 -ServerUrl https://purcheo.example.com -SumatraPath C:\Tools\SumatraPDF.exe
    .\Install-PurcheoPrintAgent.ps1 -UpdateToken github_pat_...   # only needed if the GitHub repo is private
    .\Install-PurcheoPrintAgent.ps1    # reuses the existing config.json (all options are kept)

.NOTES
    Remove with Uninstall-PurcheoPrintAgent.ps1 (add -KeepConfig to keep the token for a reinstall).

.NOTES
    Run from an elevated Windows PowerShell 5.1 prompt. The service is a small .NET Framework host,
    compiled locally with csc.exe, that runs PurcheoPrintAgent.ps1 and restarts it if it exits.
    By default the service runs as LocalSystem, which only sees printers installed machine-wide;
    pass -Credential to run as a user that has the printer installed.
#>
#Requires -Version 5.1
#Requires -RunAsAdministrator
[CmdletBinding()]
param(
    [string]$ServerUrl,
    [string]$Token,
    [string]$UpdateToken,
    [string]$SumatraPath,
    [string]$InstallDir = (Join-Path $env:ProgramData 'PurcheoPrintAgent'),
    [string]$ServiceName = 'PurcheoPrintAgent',
    [pscredential]$Credential,
    [bool]$SetPaperSize = $true
)

$ErrorActionPreference = 'Stop'

function Remove-AgentService {
    $svc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if (-not $svc) { return }
    if ($svc.Status -ne 'Stopped') { Stop-Service -Name $ServiceName -Force }
    & sc.exe delete $ServiceName | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "sc.exe delete failed ($LASTEXITCODE)" }
    # A service open in services.msc stays marked for deletion until that window is closed.
    for ($i = 0; $i -lt 10 -and (Get-Service -Name $ServiceName -ErrorAction SilentlyContinue); $i++) { Start-Sleep -Seconds 1 }
    if (Get-Service -Name $ServiceName -ErrorAction SilentlyContinue) {
        throw "Service '$ServiceName' is still marked for deletion. Close services.msc and retry."
    }
}

function Copy-IfDifferent([string]$From, [string]$To) {
    $dest = Resolve-Path -Path $To -ErrorAction SilentlyContinue
    if ($dest -and ((Resolve-Path -Path $From).Path -eq $dest.Path)) { return }
    Copy-Item -Path $From -Destination $To -Force
}

if ($PSVersionTable.PSEdition -eq 'Core') { throw 'Run this script in Windows PowerShell 5.1 (powershell.exe), not PowerShell 7.' }

$existingConfigPath = Join-Path $InstallDir 'config.json'
$existing = $null
if (Test-Path $existingConfigPath) {
    $existing = Get-Content -Raw -Path $existingConfigPath | ConvertFrom-Json
    if (-not $ServerUrl -or -not $Token) { Write-Host "Reusing settings from $existingConfigPath" }
    if (-not $ServerUrl) { $ServerUrl = [string]$existing.ServerUrl }
    if (-not $Token) { $Token = [string]$existing.Token }
}

if (-not $ServerUrl) { $ServerUrl = Read-Host 'Purcheo server URL (e.g. https://purcheo.example.com)' }
$ServerUrl = $ServerUrl.Trim().TrimEnd('/')
if ($ServerUrl -notmatch '^https?://') { throw 'ServerUrl must start with http:// or https://' }
if ($ServerUrl -notmatch '^https://') { Write-Warning 'ServerUrl is not HTTPS; the token will be sent in clear text.' }

if (-not $Token) {
    $secure = Read-Host 'Agent token' -AsSecureString
    $Token = [Runtime.InteropServices.Marshal]::PtrToStringBSTR([Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure))
}
if (-not $Token) { throw 'Token is required' }

if (-not $SumatraPath) {
    $candidates = @(
        (Join-Path $InstallDir 'SumatraPDF.exe'),
        (Join-Path $PSScriptRoot 'SumatraPDF.exe'),
        (Join-Path $env:ProgramFiles 'SumatraPDF\SumatraPDF.exe'),
        (Join-Path ${env:ProgramFiles(x86)} 'SumatraPDF\SumatraPDF.exe'),
        (Join-Path $env:LOCALAPPDATA 'SumatraPDF\SumatraPDF.exe')
    )
    $SumatraPath = $candidates | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1
}
if (-not $SumatraPath -or -not (Test-Path $SumatraPath)) {
    throw 'SumatraPDF.exe not found. Download the portable build from https://www.sumatrapdfreader.org/ and pass -SumatraPath, or place it next to this script.'
}

$csc = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
if (-not (Test-Path $csc)) { $csc = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe' }
if (-not (Test-Path $csc)) { throw '.NET Framework 4 csc.exe not found' }

$agentSource = Join-Path $PSScriptRoot 'PurcheoPrintAgent.ps1'
if (-not (Test-Path $agentSource)) { throw "PurcheoPrintAgent.ps1 not found next to this script" }

Remove-AgentService
if (-not (Test-Path $InstallDir)) { [void](New-Item -ItemType Directory -Path $InstallDir) }

# Lock the directory down first: it holds the token.
$acl = @('/inheritance:r', '/grant:r', 'SYSTEM:(OI)(CI)F', '/grant:r', 'Administrators:(OI)(CI)F')
if ($Credential) { $acl += @('/grant:r', ('{0}:(OI)(CI)RX' -f $Credential.UserName)) }
& icacls.exe $InstallDir @acl | Out-Null
if ($LASTEXITCODE -ne 0) { throw "icacls failed ($LASTEXITCODE)" }

Copy-IfDifferent $agentSource (Join-Path $InstallDir 'PurcheoPrintAgent.ps1')
Copy-IfDifferent $SumatraPath (Join-Path $InstallDir 'SumatraPDF.exe')

# Start from the existing config so options set by hand (PrintSettings, UpdateToken, ...) survive a reinstall.
$config = [ordered]@{}
if ($existing) { foreach ($p in $existing.PSObject.Properties) { $config[$p.Name] = $p.Value } }
$config.ServerUrl = $ServerUrl
$config.Token = $Token
$config.SumatraPath = (Join-Path $InstallDir 'SumatraPDF.exe')
$config.LogPath = (Join-Path $InstallDir 'agent.log')
if ($PSBoundParameters.ContainsKey('SetPaperSize') -or -not $config.Contains('SetPaperSize')) { $config.SetPaperSize = $SetPaperSize }
if ($UpdateToken) { $config.UpdateToken = $UpdateToken }
$config | ConvertTo-Json | Set-Content -Path (Join-Path $InstallDir 'config.json') -Encoding UTF8

$hostSource = @'
using System;
using System.Diagnostics;
using System.IO;
using System.ServiceProcess;
using System.Threading;

public class PurcheoPrintAgentService : ServiceBase
{
    private Process child;
    private Thread worker;
    private volatile bool stopping;

    public PurcheoPrintAgentService() { ServiceName = "PurcheoPrintAgent"; }

    public static void Main() { ServiceBase.Run(new PurcheoPrintAgentService()); }

    protected override void OnStart(string[] args)
    {
        worker = new Thread(Run);
        worker.IsBackground = true;
        worker.Start();
    }

    private void Run()
    {
        string dir = AppDomain.CurrentDomain.BaseDirectory;
        string script = Path.Combine(dir, "PurcheoPrintAgent.ps1");
        string config = Path.Combine(dir, "config.json");
        while (!stopping)
        {
            try
            {
                ProcessStartInfo psi = new ProcessStartInfo("powershell.exe",
                    "-NoProfile -NonInteractive -ExecutionPolicy Bypass -File \"" + script + "\" -ConfigPath \"" + config + "\"");
                psi.UseShellExecute = false;
                psi.CreateNoWindow = true;
                psi.WorkingDirectory = dir;
                child = Process.Start(psi);
                child.WaitForExit();
            }
            catch (Exception) { }
            if (!stopping) Thread.Sleep(5000);
        }
    }

    protected override void OnStop()
    {
        stopping = true;
        try { if (child != null && !child.HasExited) child.Kill(); } catch (Exception) { }
    }
}
'@
$csPath = Join-Path $env:TEMP 'PurcheoPrintAgentService.cs'
$exePath = Join-Path $InstallDir 'PurcheoPrintAgentService.exe'
Set-Content -Path $csPath -Value $hostSource -Encoding ASCII
try {
    & $csc /nologo /target:exe /out:$exePath /reference:System.ServiceProcess.dll $csPath
    if ($LASTEXITCODE -ne 0) { throw "csc.exe failed ($LASTEXITCODE)" }
} finally {
    Remove-Item -Path $csPath -Force -ErrorAction SilentlyContinue
}

$svcParams = @{
    Name           = $ServiceName
    BinaryPathName = "`"$exePath`""
    DisplayName    = 'Purcheo Print Agent'
    Description    = 'Polls a Purcheo server for label and thermal print jobs.'
    StartupType    = 'Automatic'
}
if ($Credential) { $svcParams.Credential = $Credential }
[void](New-Service @svcParams)
& sc.exe failure $ServiceName reset= 86400 actions= restart/5000/restart/5000/restart/30000 | Out-Null

Start-Service -Name $ServiceName
Write-Host "Service '$ServiceName' installed and started. Log: $(Join-Path $InstallDir 'agent.log')"
