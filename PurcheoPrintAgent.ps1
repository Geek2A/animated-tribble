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
$AgentVersion = '1.1.1'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$cfg = Get-Content -Raw -Path $ConfigPath | ConvertFrom-Json
$BaseUrl = ([string]$cfg.ServerUrl).TrimEnd('/')
$Token = [string]$cfg.Token
$SumatraPath = if ($cfg.SumatraPath) { [string]$cfg.SumatraPath } else { Join-Path $PSScriptRoot 'SumatraPDF.exe' }
$LogPath = if ($cfg.LogPath) { [string]$cfg.LogPath } else { Join-Path $PSScriptRoot 'agent.log' }
# Set the driver's page size from the job's media size before printing; off = use the driver default.
$SetPaperSize = if ($null -ne $cfg.SetPaperSize) { [bool]$cfg.SetPaperSize } else { $true }
$PrintSettings = if ($cfg.PrintSettings) { [string]$cfg.PrintSettings } else { 'noscale' }
# Labels are often designed a millimetre or so off the stock's catalogue size.
$PaperToleranceMm = if ($null -ne $cfg.PaperToleranceMm) { [double]$cfg.PaperToleranceMm } else { 1.5 }
$SumatraTimeoutSeconds = if ($cfg.SumatraTimeoutSeconds) { [int]$cfg.SumatraTimeoutSeconds } else { 60 }
# Labels sent straight after another can be dropped while the printer is still feeding.
$JobSettleSeconds = if ($null -ne $cfg.JobSettleSeconds) { [double]$cfg.JobSettleSeconds } else { 2 }
$PrinterReadyWaitSeconds = if ($cfg.PrinterReadyWaitSeconds) { [int]$cfg.PrinterReadyWaitSeconds } else { 15 }
$PrinterWatchIntervalSeconds = 15
# Self-update from GitHub. A private or missing repo (401/403/404) is not an error: updating pauses for UpdateBackoffHours.
$UpdateRepo = if ($cfg.UpdateRepo) { [string]$cfg.UpdateRepo } else { 'Geek2A/animated-tribble' }
$UpdateBranch = if ($cfg.UpdateBranch) { [string]$cfg.UpdateBranch } else { 'main' }
$UpdateToken = [string]$cfg.UpdateToken
$AutoUpdate = if ($null -ne $cfg.AutoUpdate) { [bool]$cfg.AutoUpdate } else { $true }
$UpdateCheckIntervalMinutes = if ($cfg.UpdateCheckIntervalMinutes) { [double]$cfg.UpdateCheckIntervalMinutes } else { 60 }
$UpdateBackoffHours = if ($cfg.UpdateBackoffHours) { [double]$cfg.UpdateBackoffHours } else { 24 }
$script:NextUpdateCheck = Get-Date

$PrinterCheckIntervalSeconds = 60
$AuthBackoffSeconds = 60
$ErrorBackoffSeconds = 10
$TempDir = Join-Path ([IO.Path]::GetTempPath()) 'PurcheoPrintAgent'
$script:PollInterval = 2
$script:DefaultPrinter = $null
$script:LastPaper = @{}
$script:PaperListed = @{}
$script:CurrentJobId = $null
$script:LastJobEnd = [datetime]::MinValue
$script:LastPrinterSummary = $null
$script:NextPrinterWatch = Get-Date
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
    if ($script:CurrentJobId) {
        $Data = if ($Data) { $Data.Clone() } else { @{} }
        if (-not $Data.ContainsKey('job_id')) { $Data['job_id'] = $script:CurrentJobId }
    }
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
            [void](Invoke-Api -Method POST -Url "$BaseUrl/api/print-agent/logs/" -Body @{ entries = $batch })
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

$PsfNs = 'http://schemas.microsoft.com/windows/2003/08/printing/printschemaframework'
$SpoolerWaitSeconds = 30

# The paper sizes the printer's driver offers: name, size and the namespace its option name needs.
function Get-PaperOptions([string]$Printer) {
    Add-Type -AssemblyName System.Printing
    $server = New-Object System.Printing.LocalPrintServer
    $queue = $null
    $reader = $null
    try {
        $queue = $server.GetPrintQueue($Printer)
        $reader = New-Object IO.StreamReader($queue.GetPrintCapabilitiesAsXml())
        $caps = New-Object Xml.XmlDocument
        $caps.LoadXml($reader.ReadToEnd())
    } finally {
        if ($reader) { $reader.Dispose() }
        if ($queue) { $queue.Dispose() }
        $server.Dispose()
    }
    $nsm = New-Object Xml.XmlNamespaceManager($caps.NameTable)
    $nsm.AddNamespace('psf', $PsfNs)
    $feature = $caps.SelectSingleNode("//psf:Feature[@name='psk:PageMediaSize']", $nsm)
    $options = @()
    if (-not $feature) { return $options }
    foreach ($o in $feature.SelectNodes('psf:Option', $nsm)) {
        $w = $o.SelectSingleNode("psf:ScoredProperty[@name='psk:MediaSizeWidth']/psf:Value", $nsm)
        $h = $o.SelectSingleNode("psf:ScoredProperty[@name='psk:MediaSizeHeight']/psf:Value", $nsm)
        if (-not $w -or -not $h) { continue }  # custom-size ranges have no fixed size
        $name = $o.GetAttribute('name')
        $prefix = if ($name.Contains(':')) { $name.Split(':')[0] } else { '' }
        $label = $o.SelectSingleNode("psf:Property[@name='psk:DisplayName']/psf:Value", $nsm)
        $options += [pscustomobject]@{
            Name      = $name
            Display   = if ($label) { $label.InnerText } else { $name }
            Prefix    = $prefix
            Namespace = if ($prefix) { $o.GetNamespaceOfPrefix($prefix) } else { '' }
            WidthMm   = [double]$w.InnerText / 1000
            HeightMm  = [double]$h.InnerText / 1000
        }
    }
    return $options
}

# Drivers list the same label either way round, so a size matches in both orientations.
function Find-PaperOption($Options, [double]$WidthMm, [double]$HeightMm) {
    $best = $null
    $bestDistance = [double]::MaxValue
    foreach ($o in $Options) {
        $d1 = [math]::Max([math]::Abs($o.WidthMm - $WidthMm), [math]::Abs($o.HeightMm - $HeightMm))
        $d2 = [math]::Max([math]::Abs($o.WidthMm - $HeightMm), [math]::Abs($o.HeightMm - $WidthMm))
        $d = [math]::Min($d1, $d2)
        if ($d -le $PaperToleranceMm -and $d -lt $bestDistance) { $best = $o; $bestDistance = $d }
    }
    return $best
}

function Set-PaperSize([string]$Printer, [double]$WidthMm, [double]$HeightMm) {
    $key = '{0}|{1}|{2}' -f $Printer, $WidthMm, $HeightMm
    if ($script:LastPaper[$Printer] -eq $key) { return }

    $options = @(Get-PaperOptions $Printer)
    if (-not $script:PaperListed[$Printer]) {
        $list = ($options | ForEach-Object { '{0} ({1:0.#}x{2:0.#} mm)' -f $_.Display, $_.WidthMm, $_.HeightMm }) -join '; '
        Write-Log "Paper sizes offered by '$Printer': $list"
        $script:PaperListed[$Printer] = $true
    }
    $match = Find-PaperOption $options $WidthMm $HeightMm
    if (-not $match) {
        throw ('The driver offers no paper size close to {0:0.##}x{1:0.##} mm' -f $WidthMm, $HeightMm)
    }

    $ticket = New-Object Xml.XmlDocument
    $ticket.LoadXml((Get-PrintConfiguration -PrinterName $Printer).PrintTicketXml)
    $nsm = New-Object Xml.XmlNamespaceManager($ticket.NameTable)
    $nsm.AddNamespace('psf', $PsfNs)
    $root = $ticket.DocumentElement
    $old = $root.SelectSingleNode("psf:Feature[@name='psk:PageMediaSize']", $nsm)
    if ($old) { [void]$root.RemoveChild($old) }

    $extraNs = if ($match.Prefix -and $match.Prefix -ne 'psk') { ' xmlns:{0}="{1}"' -f $match.Prefix, $match.Namespace } else { '' }
    $frag = $ticket.CreateDocumentFragment()
    $frag.InnerXml = '<psf:Feature xmlns:psf="{0}" xmlns:psk="http://schemas.microsoft.com/windows/2003/08/printing/printschemakeywords"{1} name="psk:PageMediaSize"><psf:Option name="{2}" /></psf:Feature>' -f $PsfNs, $extraNs, $match.Name
    [void]$root.AppendChild($frag)
    Set-PrintConfiguration -PrinterName $Printer -PrintTicketXml $ticket.OuterXml
    $script:LastPaper[$Printer] = $key
    Write-Log ("Paper size for '{0}' set to {1} ({2:0.#}x{3:0.#} mm)" -f $Printer, $match.Display, $match.WidthMm, $match.HeightMm)
}

function Get-ErrorDetail($ErrorRecord) {
    $parts = @($ErrorRecord.Exception.Message)
    $inner = $ErrorRecord.Exception.InnerException
    while ($inner) { $parts += $inner.Message; $inner = $inner.InnerException }
    if ($ErrorRecord.FullyQualifiedErrorId) { $parts += "[$($ErrorRecord.FullyQualifiedErrorId)]" }
    return ($parts | Select-Object -Unique) -join ' | '
}

# SumatraPDF exits 0 once the job is handed to Windows, so check that the spooler really took it.
function Wait-SpoolerDone([string]$Printer, [datetime]$Since) {
    $deadline = (Get-Date).AddSeconds($SpoolerWaitSeconds)
    $seen = $false
    $last = $null
    $watch = [Diagnostics.Stopwatch]::StartNew()
    while ((Get-Date) -lt $deadline) {
        try {
            $jobs = @(Get-PrintJob -PrinterName $Printer -ErrorAction Stop |
                Where-Object { $_.SubmittedTime -ge $Since.AddSeconds(-5) })
        } catch {
            Write-Log "Could not read the Windows print queue: $($_.Exception.Message)" 'WARNING'
            return
        }
        $failed = $jobs | Where-Object { "$($_.JobStatus)" -match 'Error|Offline|PaperOut|Blocked|UserIntervention' } | Select-Object -First 1
        if ($failed) { throw "Windows print queue reports '$($failed.JobStatus)' for the job" }
        if ($jobs.Count -eq 0) {
            # A label can print and leave the queue before the first look.
            if (-not $seen) {
                Write-Log "No job left in the Windows print queue for '$Printer' (already printed, or never queued)"
            } else {
                Write-Log ("Windows print queue: job left the queue after {0:0.0} s ({1} bytes, {2} page(s), last status {3})" -f $watch.Elapsed.TotalSeconds, $last.Size, $last.TotalPages, $last.JobStatus)
            }
            return
        }
        $seen = $true
        $last = $jobs | Select-Object -First 1
        Start-Sleep -Milliseconds 500
    }
    $status = (Get-Printer -Name $Printer -ErrorAction SilentlyContinue).PrinterStatus
    throw "The job was still in the Windows print queue after $SpoolerWaitSeconds s (printer status: $status)"
}

$DetectedErrorNames = @{
    0 = 'unknown'; 1 = 'other'; 2 = 'no error'; 3 = 'low paper'; 4 = 'no paper'; 5 = 'low toner'
    6 = 'no toner'; 7 = 'door open'; 8 = 'jammed'; 9 = 'offline'; 10 = 'service requested'; 11 = 'output bin full'
}
$BlockingStatus = 'Error|Offline|PaperOut|PaperJam|PaperProblem|DoorOpen|UserIntervention|NotAvailable|NoToner|OutOfMemory'
# Offline can flicker on a printer that has just gone idle, so only these fail a job after the fact.
$HardStatus = 'Error|PaperOut|PaperJam|DoorOpen|UserIntervention|NoToner'
$BusyStatus = 'Printing|Busy|Processing|IOActive|Initializing|WarmingUp'

# What Windows knows about the printer's physical state.
function Get-PrinterState([string]$Printer) {
    $p = Get-Printer -Name $Printer -ErrorAction SilentlyContinue
    $cim = Get-CimInstance -ClassName Win32_Printer -Filter ("Name='{0}'" -f $Printer.Replace('\', '\\').Replace("'", "''")) -ErrorAction SilentlyContinue
    $status = if ($p) { [string]$p.PrinterStatus } else { 'NotFound' }
    $detected = $null
    if ($cim -and $null -ne $cim.DetectedErrorState) { $detected = $DetectedErrorNames[[int]$cim.DetectedErrorState] }
    $offline = [bool]($cim -and $cim.WorkOffline)
    $summary = "status $status"
    if ($offline) { $summary += ', set to work offline' }
    if ($detected -and $detected -ne 'no error' -and $detected -ne 'unknown') { $summary += ", $detected" }
    $blocked = $offline -or ($status -match $BlockingStatus) -or ($detected -in @('no paper', 'door open', 'jammed', 'offline'))
    $hard = ($status -match $HardStatus) -or ($detected -in @('no paper', 'door open', 'jammed'))
    [pscustomobject]@{ Status = $status; Summary = $summary; Blocked = $blocked; Hard = $hard; Busy = ($status -match $BusyStatus) }
}

# Waits until the printer is idle and healthy, and long enough after the previous label.
function Wait-PrinterReady([string]$Printer) {
    $settle = $script:LastJobEnd.AddSeconds($JobSettleSeconds) - (Get-Date)
    if ($settle.TotalMilliseconds -gt 0) { Start-Sleep -Milliseconds ([int]$settle.TotalMilliseconds) }
    $deadline = (Get-Date).AddSeconds($PrinterReadyWaitSeconds)
    $busyDeadline = (Get-Date).AddSeconds(5)
    $state = Get-PrinterState $Printer
    while (($state.Blocked -and (Get-Date) -lt $deadline) -or ($state.Busy -and (Get-Date) -lt $busyDeadline)) {
        Start-Sleep -Seconds 1
        $state = Get-PrinterState $Printer
    }
    Write-Log "Printer '$Printer' before the job: $($state.Summary)"
    if ($state.Blocked) { throw "Printer '$Printer' is not ready: $($state.Summary)" }
}

# A label that never comes out often leaves the printer in an error state a moment later.
function Confirm-PrinterAfter([string]$Printer) {
    Start-Sleep -Seconds 2
    $state = Get-PrinterState $Printer
    Write-Log "Printer '$Printer' after the job: $($state.Summary)" $(if ($state.Blocked) { 'WARNING' } else { 'INFO' })
    if ($state.Hard) { throw "Printer '$Printer' reports a problem after printing: $($state.Summary)" }
}

# Logs only changes, so a printer that goes offline or runs out of labels shows up with its time.
function Watch-PrinterState {
    if ((Get-Date) -lt $script:NextPrinterWatch -or -not $script:DefaultPrinter) { return }
    $script:NextPrinterWatch = (Get-Date).AddSeconds($PrinterWatchIntervalSeconds)
    $state = Get-PrinterState $script:DefaultPrinter
    if ($state.Summary -eq $script:LastPrinterSummary) { return }
    $level = if ($state.Blocked) { 'WARNING' } else { 'INFO' }
    $was = if ($script:LastPrinterSummary) { $script:LastPrinterSummary } else { 'not checked yet' }
    Write-Log "Printer '$($script:DefaultPrinter)' state changed: $was -> $($state.Summary)" $level
    $script:LastPrinterSummary = $state.Summary
}

# Windows' own print service errors since the job started (the Admin log is on by default).
function Write-PrintServiceErrors([datetime]$Since) {
    try {
        $events = @(Get-WinEvent -FilterHashtable @{ LogName = 'Microsoft-Windows-PrintService/Admin'; StartTime = $Since.AddSeconds(-2) } -ErrorAction Stop |
            Where-Object { $_.Level -le 3 } | Select-Object -First 5)
    } catch { return }  # no events (or no access) is the normal case
    foreach ($e in $events) {
        $text = ([string]$e.Message) -replace '\s+', ' '
        if ($text.Length -gt 300) { $text = $text.Substring(0, 300) }
        Write-Log "Windows print service event $($e.Id): $text" 'WARNING'
    }
}

function Test-DownloadedDocument([string]$Path, $Job) {
    $actual = (Get-Item $Path).Length
    if ($Job.size_bytes -and $actual -ne [int64]$Job.size_bytes) {
        throw "Downloaded document is $actual bytes but the server sent $($Job.size_bytes)"
    }
    if ($Job.sha256) {
        $hash = (Get-FileHash -Path $Path -Algorithm SHA256).Hash.ToLowerInvariant()
        if ($hash -ne ([string]$Job.sha256).ToLowerInvariant()) { throw 'Downloaded document does not match the server checksum' }
    }
    Write-Log "Downloaded the document: $actual bytes$(if ($Job.sha256) { ', checksum matches' })"
}

function Invoke-PrintJob($Job) {
    $printer = [string]$Job.printer_name
    if (-not $printer) { $printer = $script:DefaultPrinter }
    if (-not $printer) { throw 'No printer configured for this job or agent' }
    if (-not (Get-Printer -Name $printer -ErrorAction SilentlyContinue)) {
        throw "Printer '$printer' not found on this machine"
    }
    if (-not (Test-Path $SumatraPath)) { throw "SumatraPDF not found at $SumatraPath" }
    Wait-PrinterReady $printer

    if (-not (Test-Path $TempDir)) { [void](New-Item -ItemType Directory -Path $TempDir) }
    $pdf = Join-Path $TempDir ("job-{0}.pdf" -f $Job.id)
    try {
        [void](Invoke-Api -Method GET -Url (Resolve-PdfUrl ([string]$Job.pdf_path)) -OutFile $pdf)
        Test-DownloadedDocument $pdf $Job

        if ($SetPaperSize -and $Job.media_width_mm -and $Job.media_height_mm) {
            try { Set-PaperSize $printer ([double]$Job.media_width_mm) ([double]$Job.media_height_mm) }
            catch { Write-Log "Could not set paper size, using driver default: $(Get-ErrorDetail $_)" 'WARNING' }
        }

        $started = Get-Date
        $args = @('-print-to', $printer, '-print-settings', $PrintSettings, '-silent', '-exit-when-done', $pdf)
        # No -Wait: it also waits for every process Sumatra starts, which can hold the job for minutes.
        $p = Start-Process -FilePath $SumatraPath -ArgumentList ($args | ForEach-Object { '"{0}"' -f $_ }) -PassThru -WindowStyle Hidden
        $null = $p.Handle  # keeps the exit code readable after the process ends
        if (-not $p.WaitForExit($SumatraTimeoutSeconds * 1000)) {
            try { $p.Kill() } catch { }
            throw "SumatraPDF did not finish within $SumatraTimeoutSeconds s and was stopped"
        }
        Write-Log ("SumatraPDF exited with code {0} after {1:0.0} s for printer '{2}' (settings: {3})" -f $p.ExitCode, ((Get-Date) - $started).TotalSeconds, $printer, $PrintSettings)
        if ($p.ExitCode -ne 0) { throw "SumatraPDF exited with code $($p.ExitCode)" }
        try { Wait-SpoolerDone $printer $started } finally { Write-PrintServiceErrors $started }
        Confirm-PrinterAfter $printer
    } finally {
        $script:LastJobEnd = Get-Date
        Remove-Item -Path $pdf -Force -ErrorAction SilentlyContinue
    }
}

# Replaces this script with the branch's copy when it differs; the service host restarts the agent afterwards.
function Update-AgentIfNewer {
    if (-not (Test-Path $TempDir)) { [void](New-Item -ItemType Directory -Path $TempDir) }
    $new = Join-Path $TempDir 'PurcheoPrintAgent.ps1.new'
    $headers = @{ Accept = 'application/vnd.github.raw'; 'User-Agent' = 'PurcheoPrintAgent' }
    if ($UpdateToken) { $headers.Authorization = "Bearer $UpdateToken" }
    try {
        try {
            Invoke-WebRequest -Uri "https://api.github.com/repos/$UpdateRepo/contents/PurcheoPrintAgent.ps1?ref=$UpdateBranch" `
                -Headers $headers -OutFile $new -UseBasicParsing -TimeoutSec 60
        } catch {
            $status = if ($_.Exception.Response) { [int]$_.Exception.Response.StatusCode } else { 0 }
            if ($status -in 401, 403, 404) {
                $script:NextUpdateCheck = (Get-Date).AddHours($UpdateBackoffHours)
                return $false
            }
            throw
        }
        if ((Get-Item $new).Length -eq 0) { throw 'downloaded file is empty' }
        if ((Get-FileHash -Path $new -Algorithm SHA256).Hash -eq (Get-FileHash -Path $PSCommandPath -Algorithm SHA256).Hash) { return $false }

        # Never install a copy that does not parse: the service would restart into it forever.
        $parseErrors = $null
        [void][Management.Automation.Language.Parser]::ParseFile($new, [ref]$null, [ref]$parseErrors)
        if ($parseErrors.Count -gt 0) { throw "downloaded script has syntax errors: $($parseErrors[0].Message)" }

        Copy-Item -Path $PSCommandPath -Destination "$PSCommandPath.bak" -Force
        Copy-Item -Path $new -Destination $PSCommandPath -Force
        Write-Log "Updated agent from $UpdateRepo@$UpdateBranch (was $AgentVersion); restarting"
        return $true
    } finally {
        Remove-Item -Path $new -Force -ErrorAction SilentlyContinue
    }
}

function Send-JobResult([string]$JobId, [bool]$Success, [string]$ErrorText) {
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
    if ($AutoUpdate -and (Get-Date) -ge $script:NextUpdateCheck) {
        $script:NextUpdateCheck = (Get-Date).AddMinutes($UpdateCheckIntervalMinutes)
        try {
            if (Update-AgentIfNewer) {
                try { Send-PendingLogs } catch { }
                exit 0
            }
        } catch {
            Write-Log "Update check failed: $(Get-ErrorDetail $_)" 'WARNING' -NoShip
        }
    }
    try {
        if (((Get-Date) - $lastPrinterCheck).TotalSeconds -ge $PrinterCheckIntervalSeconds) {
            Send-HeartbeatIfChanged
            $lastPrinterCheck = Get-Date
        }
        Watch-PrinterState

        $r = Invoke-Api -Method GET -Url "$BaseUrl/api/print-agent/jobs/next/"
        if ($r.Status -ne 204 -and $r.Json -and $r.Json.job) {
            $job = $r.Json.job
            # Job ids are UUIDs; parsing also keeps odd values out of file names and URLs.
            $job.id = ([guid][string]$job.id).ToString()
            $script:CurrentJobId = $job.id
            Write-Log "Job $($job.id): '$($job.title)' $($job.pages) page(s) $($job.media_width_mm)x$($job.media_height_mm)mm" 'INFO' @{ job_id = $job.id }
            try {
                Invoke-PrintJob $job
                Send-JobResult $job.id $true $null
                Write-Log "Job $($job.id) printed" 'INFO' @{ job_id = $job.id }
            } catch {
                Write-Log "Job $($job.id) failed: $($_.Exception.Message)" 'ERROR' @{ job_id = $job.id }
                Send-JobResult $job.id $false $_.Exception.Message
            } finally {
                $script:CurrentJobId = $null
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
