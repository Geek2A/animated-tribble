<#
.SYNOPSIS
    Purcheo print agent: polls a Purcheo server for label/thermal print jobs and prints them.

.DESCRIPTION
    Contract (all requests carry "Authorization: Bearer <token>"):
      POST /api/print-agent/heartbeat/          {"version","printers":[...]} -> {"name","printer_name","poll_interval_seconds"}
      GET  /api/print-agent/jobs/next/          -> {"job":{...}} or 204
      GET  <job.pdf_path>                       -> PDF, one page per label, already sized and rotated
      POST /api/print-agent/jobs/<id>/result/   {"success":true} | {"success":false,"error":"..."}

    Runs standalone for testing, or as a Windows service via Install-PurcheoPrintAgent.ps1.
    Requires Windows PowerShell 5.1+ and SumatraPDF for silent PDF printing.
#>
[CmdletBinding()]
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot 'config.json')
)

$ErrorActionPreference = 'Stop'
$AgentVersion = '1.0.0'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$cfg = Get-Content -Raw -Path $ConfigPath | ConvertFrom-Json
$BaseUrl = ([string]$cfg.ServerUrl).TrimEnd('/')
$Token = [string]$cfg.Token
$SumatraPath = if ($cfg.SumatraPath) { [string]$cfg.SumatraPath } else { Join-Path $PSScriptRoot 'SumatraPDF.exe' }
$LogPath = if ($cfg.LogPath) { [string]$cfg.LogPath } else { Join-Path $PSScriptRoot 'agent.log' }
# Set the driver's page size from the job's media size before printing; off = use the driver default.
$SetPaperSize = if ($null -ne $cfg.SetPaperSize) { [bool]$cfg.SetPaperSize } else { $true }
$PrintSettings = if ($cfg.PrintSettings) { [string]$cfg.PrintSettings } else { 'noscale' }

$PrinterCheckIntervalSeconds = 60
$AuthBackoffSeconds = 60
$ErrorBackoffSeconds = 10
$TempDir = Join-Path ([IO.Path]::GetTempPath()) 'PurcheoPrintAgent'
$script:PollInterval = 2
$script:DefaultPrinter = $null
$script:LastPaper = @{}
$script:SentPrinters = $null
# Lines wait here until the server accepts them, so an outage is backfilled with original timestamps.
$LogBatchSize = 200
$LogQueueLimit = 5000
$script:LogQueue = New-Object 'System.Collections.Generic.List[object]'

function Write-Log([string]$Message, [string]$Level = 'INFO', [hashtable]$Data = $null, [switch]$NoShip) {
    if ($Level -eq 'WARN') { $Level = 'WARNING' }
    $now = Get-Date
    $line = '{0} [{1}] {2}' -f $now.ToString('yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Write-Host $line
    try {
        if ((Test-Path $LogPath) -and (Get-Item $LogPath).Length -gt 5MB) {
            Move-Item -Path $LogPath -Destination "$LogPath.1" -Force
        }
        Add-Content -Path $LogPath -Value $line -Encoding UTF8
    } catch { }
    if ($NoShip) { return }
    $text = if ($Message) { $Message } else { '(empty)' }
    if ($text.Length -gt 4000) { $text = $text.Substring(0, 4000) }
    $entry = [ordered]@{
        id      = [guid]::NewGuid().ToString()
        time    = $now.ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss.fff'Z'", [Globalization.CultureInfo]::InvariantCulture)
        level   = $Level
        message = $text
    }
    if ($Data) { $entry.data = $Data }
    $script:LogQueue.Add($entry)
    if ($script:LogQueue.Count -gt $LogQueueLimit) {
        $script:LogQueue.RemoveRange(0, $script:LogQueue.Count - $LogQueueLimit)
    }
}

function Invoke-Api {
    param(
        [string]$Method,
        [string]$Url,
        $Body = $null,
        [string]$OutFile = $null
    )
    $params = @{
        Method          = $Method
        Uri             = $Url
        Headers         = @{ Authorization = "Bearer $Token"; Accept = 'application/json' }
        UseBasicParsing = $true
        TimeoutSec      = 30
    }
    if ($null -ne $Body) {
        $params.Body = [Text.Encoding]::UTF8.GetBytes(($Body | ConvertTo-Json -Depth 5 -Compress))
        $params.ContentType = 'application/json; charset=utf-8'
    }
    if ($OutFile) { $params.OutFile = $OutFile }
    try {
        $resp = Invoke-WebRequest @params
    } catch {
        $status = 0
        if ($_.Exception.Response) { $status = [int]$_.Exception.Response.StatusCode }
        $err = New-Object System.Exception("HTTP $status from $Method $Url : $($_.Exception.Message)")
        $err.Data['Status'] = $status
        throw $err
    }
    $result = @{ Status = [int]$resp.StatusCode; Json = $null }
    if (-not $OutFile -and $result.Status -ne 204 -and $resp.RawContentStream) {
        $text = [Text.Encoding]::UTF8.GetString($resp.RawContentStream.ToArray())
        if ($text.Trim()) { $result.Json = $text | ConvertFrom-Json }
    }
    return $result
}

function Get-ErrorStatus($ErrorRecord) {
    $ex = $ErrorRecord.Exception
    if ($ex.Data.Contains('Status')) { return [int]$ex.Data['Status'] }
    return 0
}

# Sent once at startup, then only when the Windows printer list changes.
function Send-HeartbeatIfChanged {
    $printers = @(Get-Printer | ForEach-Object { $_.Name } | Sort-Object)
    $signature = $printers -join "`n"
    if ($null -ne $script:SentPrinters -and $signature -eq $script:SentPrinters) { return }
    $r = Invoke-Api -Method POST -Url "$BaseUrl/api/print-agent/heartbeat/" -Body @{
        version  = $AgentVersion
        printers = $printers
    }
    if ($r.Json) {
        if ($r.Json.poll_interval_seconds -and [double]$r.Json.poll_interval_seconds -gt 0) {
            $script:PollInterval = [double]$r.Json.poll_interval_seconds
        }
        $script:DefaultPrinter = [string]$r.Json.printer_name
    }
    $script:SentPrinters = $signature
    Write-Log "Heartbeat sent ($($printers.Count) printer(s))"
}

function Send-PendingLogs {
    $batches = 0
    while ($script:LogQueue.Count -gt 0 -and $batches -lt 10) {
        $n = [math]::Min($LogBatchSize, $script:LogQueue.Count)
        $batch = $script:LogQueue.GetRange(0, $n).ToArray()
        try {
            [void](Invoke-Api -Method POST -Url "$BaseUrl/api/print-agent/logs/" -Body @{ lines = $batch })
        } catch {
            if ((Get-ErrorStatus $_) -ne 400) { throw }
            # A batch the server rejects outright would otherwise block the queue forever.
            Write-Log "Server rejected a log batch of $n line(s); dropping it: $($_.Exception.Message)" 'WARNING' -NoShip
        }
        $script:LogQueue.RemoveRange(0, $n)
        $batches++
    }
}

function Resolve-PdfUrl([string]$PdfPath) {
    $uri = New-Object Uri((New-Object Uri("$BaseUrl/")), $PdfPath)
    $base = New-Object Uri($BaseUrl)
    # Never send the bearer token to a host other than the configured server.
    if ($uri.Scheme -ne $base.Scheme -or $uri.Authority -ne $base.Authority) {
        throw "pdf_path points to a different host: $uri"
    }
    return $uri.AbsoluteUri
}

function Set-PaperSize([string]$Printer, [double]$WidthMm, [double]$HeightMm) {
    $key = '{0}|{1}|{2}' -f $Printer, $WidthMm, $HeightMm
    if ($script:LastPaper[$Printer] -eq $key) { return }

    $ptNs = 'http://schemas.microsoft.com/windows/2003/08/printing/printticket'
    $widthUm = [int][math]::Round($WidthMm * 1000)
    $heightUm = [int][math]::Round($HeightMm * 1000)

    $ticket = [xml](Get-PrintConfiguration -PrinterName $Printer).PrintTicketXml
    $nsm = New-Object Xml.XmlNamespaceManager($ticket.NameTable)
    $nsm.AddNamespace('pt', $ptNs)
    $root = $ticket.DocumentElement
    $old = $root.SelectSingleNode("pt:Feature[@name='psk:PageMediaSize']", $nsm)
    if ($old) { [void]$root.RemoveChild($old) }

    $frag = $ticket.CreateDocumentFragment()
    $frag.InnerXml = @"
<Feature xmlns="$ptNs" xmlns:psk="http://schemas.microsoft.com/windows/2003/08/printing/printschemakeywords" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" name="psk:PageMediaSize">
  <Option name="psk:CustomMediaSize">
    <ScoredProperty name="psk:MediaSizeWidth"><Value xsi:type="xsd:integer">$widthUm</Value></ScoredProperty>
    <ScoredProperty name="psk:MediaSizeHeight"><Value xsi:type="xsd:integer">$heightUm</Value></ScoredProperty>
  </Option>
</Feature>
"@
    [void]$root.AppendChild($frag)
    Set-PrintConfiguration -PrinterName $Printer -PrintTicketXml $ticket.OuterXml
    $script:LastPaper[$Printer] = $key
    Write-Log "Paper size for '$Printer' set to $WidthMm x $HeightMm mm"
}

function Invoke-PrintJob($Job) {
    $printer = [string]$Job.printer_name
    if (-not $printer) { $printer = $script:DefaultPrinter }
    if (-not $printer) { throw 'No printer configured for this job or agent' }
    if (-not (Get-Printer -Name $printer -ErrorAction SilentlyContinue)) {
        throw "Printer '$printer' not found on this machine"
    }
    if (-not (Test-Path $SumatraPath)) { throw "SumatraPDF not found at $SumatraPath" }

    if (-not (Test-Path $TempDir)) { [void](New-Item -ItemType Directory -Path $TempDir) }
    $pdf = Join-Path $TempDir ("job-{0}.pdf" -f ([int]$Job.id))
    try {
        [void](Invoke-Api -Method GET -Url (Resolve-PdfUrl ([string]$Job.pdf_path)) -OutFile $pdf)

        if ($SetPaperSize -and $Job.media_width_mm -and $Job.media_height_mm) {
            try { Set-PaperSize $printer ([double]$Job.media_width_mm) ([double]$Job.media_height_mm) }
            catch { Write-Log "Could not set paper size, using driver default: $($_.Exception.Message)" 'WARN' }
        }

        $args = @('-print-to', $printer, '-print-settings', $PrintSettings, '-silent', '-exit-when-done', $pdf)
        $p = Start-Process -FilePath $SumatraPath -ArgumentList ($args | ForEach-Object { '"{0}"' -f $_ }) -Wait -PassThru -WindowStyle Hidden
        if ($p.ExitCode -ne 0) { throw "SumatraPDF exited with code $($p.ExitCode)" }
    } finally {
        Remove-Item -Path $pdf -Force -ErrorAction SilentlyContinue
    }
}

function Send-JobResult([int]$JobId, [bool]$Success, [string]$ErrorText) {
    $body = @{ success = $Success }
    if (-not $Success) { $body.error = $ErrorText }
    for ($i = 1; $i -le 3; $i++) {
        try {
            [void](Invoke-Api -Method POST -Url "$BaseUrl/api/print-agent/jobs/$JobId/result/" -Body $body)
            return
        } catch {
            Write-Log "Reporting result for job $JobId failed (attempt $i): $($_.Exception.Message)" 'WARN'
            Start-Sleep -Seconds 2
        }
    }
}

Write-Log "Purcheo print agent $AgentVersion starting, server $BaseUrl"
$lastPrinterCheck = [datetime]::MinValue

while ($true) {
    $delay = $script:PollInterval
    try {
        if (((Get-Date) - $lastPrinterCheck).TotalSeconds -ge $PrinterCheckIntervalSeconds) {
            Send-HeartbeatIfChanged
            $lastPrinterCheck = Get-Date
        }

        $r = Invoke-Api -Method GET -Url "$BaseUrl/api/print-agent/jobs/next/"
        if ($r.Status -ne 204 -and $r.Json -and $r.Json.job) {
            $job = $r.Json.job
            Write-Log "Job $($job.id): '$($job.title)' $($job.pages) page(s) $($job.media_width_mm)x$($job.media_height_mm)mm" 'INFO' @{ job_id = [int]$job.id }
            try {
                Invoke-PrintJob $job
                Send-JobResult ([int]$job.id) $true $null
                Write-Log "Job $($job.id) printed" 'INFO' @{ job_id = [int]$job.id }
            } catch {
                Write-Log "Job $($job.id) failed: $($_.Exception.Message)" 'ERROR' @{ job_id = [int]$job.id }
                Send-JobResult ([int]$job.id) $false $_.Exception.Message
            }
            $delay = 0  # check for another job immediately
        }
        Send-PendingLogs
    } catch {
        $status = Get-ErrorStatus $_
        if ($status -eq 401) {
            Write-Log 'Unauthorized: token is wrong or the agent is disabled' 'ERROR'
            $delay = $AuthBackoffSeconds
        } else {
            Write-Log $_.Exception.Message 'ERROR'
            $lastPrinterCheck = [datetime]::MinValue
            $delay = $ErrorBackoffSeconds
        }
    }
    if ($delay -gt 0) { Start-Sleep -Milliseconds ([int]($delay * 1000)) }
}
