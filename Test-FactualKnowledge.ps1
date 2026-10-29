#requires -Version 5.1
<#
.SYNOPSIS
Benchmarks five factual knowledge prompts against five local Ollama models.
.DESCRIPTION
Windows PowerShell 5.1 or PowerShell 7 on Windows. No administrator rights,
external modules, Python, or Ollama CLI parsing are required. Ollama must be
running and all requested models must already be downloaded.

Requests run sequentially. An independent PowerShell runspace performs HTTP
requests while the main thread samples total system CPU and physical RAM.
HTTP timing is measured inside that worker, independently of sampling delays.

Each prompt is a fresh request: previous answers/context are never passed back.
Preloading is timed separately; prompt results therefore describe warm loads,
not cold starts. There is no additional unscored generation warm-up.

CPU percentages use deltas of Windows raw processor counters (all logical CPUs,
normalized 0-100). RAM is (visible physical RAM - free physical RAM) / total.
Both include Windows, other apps, Ollama, and this measurement script. They do
not represent Ollama-only usage. Peaks are sampled peaks, not continuous maxima.

Default output cap is 2048 tokens; answers hitting it are marked Truncated.
Use -NumPredict -1 for no explicit output token cap (timeout still applies).
The complete returned text and JSON are saved, including any thinking field.
No thinking override is sent: installed model defaults apply and are captured
in model-info.json. Check these before interpreting cross-model comparisons.

Output: one new timestamped directory, measurements.csv (one row per prompt),
resource-samples.csv, run-info.json, model metadata, requests, text and JSON.
CSV rows are appended after each prompt, so earlier results survive later errors.
Do not run other inference clients during this benchmark.
.EXAMPLE
.\Test-FactualKnowledge.ps1
.EXAMPLE
.\Test-FactualKnowledge.ps1 -OutputDirectory C:\LLM-Benchmark -NumPredict -1
#>
[CmdletBinding()]
param(
    [string[]]$Models = @('gemma4:e4b', 'gemma4:e2b', 'llama3.2:3b', 'llama3.2:1b', 'phi4-mini:3.8b'),
    [ValidateRange(1024, 131072)][int]$ContextWindow = 4096,
    [ValidateScript({ $_ -eq -1 -or $_ -gt 0 })][int]$NumPredict = 2048,
    [ValidateRange(250, 10000)][int]$SampleIntervalMs = 1000,
    [ValidateRange(30, 86400)][int]$RequestTimeoutSeconds = 1800,
    [string]$OutputDirectory = (Join-Path $PSScriptRoot 'BenchmarkResults')
)

$ErrorActionPreference = 'Stop'
if ($env:OS -ne 'Windows_NT') { throw 'Run this script on the Windows test laptop.' }
if ($Models.Count -eq 0) { throw 'At least one model is required.' }
$BaseUrl = 'http://127.0.0.1:11434'
$Utf8 = New-Object System.Text.UTF8Encoding($false)

# These are the five original prompts, unchanged. Expected answers are not sent.
$Prompts = @(
    [pscustomobject]@{ Id = 'F01'; Title = 'HTTP Status Codes'; Text = 'Explain the differences between HTTP status codes 401, 403, and 404. For each status code, describe what it means and provide one example of when a web application would return it.' },
    [pscustomobject]@{ Id = 'F02'; Title = 'Computer Hardware'; Text = 'Explain the differences between RAM, CPU cache, and SSD storage. Describe the primary purpose of each component and whether its contents are typically retained when a computer is powered off.' },
    [pscustomobject]@{ Id = 'F03'; Title = 'Networking Fundamentals'; Text = 'Describe the primary purposes of DNS, DHCP, and NAT in a typical enterprise network. Explain how each technology contributes to network communication and give one practical example of its use.' },
    [pscustomobject]@{ Id = 'F04'; Title = 'Encryption and Hashing'; Text = 'Explain the differences between symmetric encryption, asymmetric encryption, and cryptographic hashing. For each, describe its primary purpose, whether the operation is reversible, and provide an example of an algorithm commonly used for it.' },
    [pscustomobject]@{ Id = 'F05'; Title = 'Web Authentication and Sessions'; Text = 'Explain the differences between session-based authentication and token-based authentication using JSON Web Tokens (JWTs). Describe how each approach typically maintains authentication state, how logout or session invalidation can be implemented, and one security concern associated with each approach.' }
)

function Write-JsonFile {
    param($Value, [string]$Path)
    [IO.File]::WriteAllText($Path, ($Value | ConvertTo-Json -Depth 30), $Utf8)
}

function Get-SystemSnapshot {
    # CIM class/property names work independently of localized PerfMon names.
    $cpu = Get-CimInstance Win32_PerfRawData_PerfOS_Processor -Filter "Name='_Total'"
    $os = Get-CimInstance Win32_OperatingSystem
    if (-not $cpu -or [double]$os.TotalVisibleMemorySize -le 0) {
        throw 'Windows CPU or memory counters are unavailable.'
    }
    [pscustomobject]@{
        TimestampUtc = [DateTime]::UtcNow.ToString('o')
        IdleTicks = [decimal]$cpu.PercentProcessorTime
        ClockTicks = [decimal]$cpu.Timestamp_Sys100NS
        MemoryUsedPercent = 100.0 * (1.0 - [double]$os.FreePhysicalMemory / [double]$os.TotalVisibleMemorySize)
    }
}

function Invoke-ControlRequest {
    param([string]$Path, $Body)
    $requestArgs = @{ Uri = "$BaseUrl$Path"; TimeoutSec = $RequestTimeoutSeconds; ErrorAction = 'Stop' }
    if ($null -ne $Body) {
        $requestArgs.Method = 'Post'
        $requestArgs.ContentType = 'application/json; charset=utf-8'
        $requestArgs.Body = [Text.Encoding]::UTF8.GetBytes(($Body | ConvertTo-Json -Depth 10))
    }
    Invoke-RestMethod @requestArgs
}

function Stop-BenchmarkModel {
    param([string]$Model)
    $null = Invoke-ControlRequest '/api/generate' @{ model = $Model; keep_alive = 0; stream = $false }
    $wait = [Diagnostics.Stopwatch]::StartNew()
    do {
        $running = Invoke-ControlRequest '/api/ps' $null
        if (@($running.models).Count -eq 0) { return }
        Start-Sleep -Milliseconds 250
    } while ($wait.Elapsed.TotalSeconds -lt 30)
    throw 'Models remain loaded after unloading. Benchmark stopped to avoid overlapping models.'
}

# The worker measures the actual HTTP round trip, including receiving the body.
# It returns errors as data so a failed request can still produce a CSV row.
$HttpWorker = {
    param([string]$Uri, [string]$Json, [int]$Timeout)
    $ErrorActionPreference = 'Stop'
    Add-Type -AssemblyName System.Net.Http
    $client = New-Object System.Net.Http.HttpClient
    $client.Timeout = [TimeSpan]::FromSeconds($Timeout)
    $content = New-Object System.Net.Http.StringContent($Json, [Text.Encoding]::UTF8, 'application/json')
    $response = $null
    $raw = $null
    $errorText = $null
    $httpStatus = $null
    $started = [DateTime]::UtcNow.ToString('o')
    $timer = [Diagnostics.Stopwatch]::StartNew()
    try {
        $response = $client.PostAsync($Uri, $content).GetAwaiter().GetResult()
        $raw = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
        $httpStatus = [int]$response.StatusCode
        if (-not $response.IsSuccessStatusCode) { $errorText = "HTTP $httpStatus : $raw" }
    }
    catch { $errorText = $_.Exception.ToString() }
    finally {
        $timer.Stop()
        if ($null -ne $response) { $response.Dispose() }
        $content.Dispose()
        $client.Dispose()
    }
    [pscustomobject]@{
        StartedUtc = $started; EndToEndSeconds = $timer.Elapsed.TotalSeconds
        RawBody = $raw; ErrorText = $errorText; HttpStatus = $httpStatus
    }
}

function Invoke-MeasuredPrompt {
    param($Body, [string]$Model, [string]$PromptId)
    $worker = [PowerShell]::Create()
    $samples = New-Object 'System.Collections.Generic.List[object]'
    $sampleErrors = New-Object 'System.Collections.Generic.List[string]'
    try {
        $null = $worker.AddScript($HttpWorker.ToString()).AddArgument("$BaseUrl/api/generate").AddArgument(($Body | ConvertTo-Json -Depth 10)).AddArgument($RequestTimeoutSeconds)
        $previous = Get-SystemSnapshot
        $handle = $worker.BeginInvoke()
        try {
            # WaitOne wakes immediately when HTTP finishes; timing never includes
            # an artificial one-second delay after a short response.
            do {
                $finished = $handle.AsyncWaitHandle.WaitOne($SampleIntervalMs)
                try {
                    $current = Get-SystemSnapshot
                    $ticks = $current.ClockTicks - $previous.ClockTicks
                    if ($ticks -gt 0) {
                        $cpuPercent = 100.0 * (1.0 - [double](($current.IdleTicks - $previous.IdleTicks) / $ticks))
                        $samples.Add([pscustomobject]@{
                            Model = $Model; PromptId = $PromptId; TimestampUtc = $current.TimestampUtc
                            IntervalSeconds = [double]$ticks / 1e7
                            CpuPercent = [Math]::Max(0, [Math]::Min(100, $cpuPercent))
                            SystemMemoryPercent = $current.MemoryUsedPercent
                        })
                    }
                    $previous = $current
                }
                catch { $sampleErrors.Add($_.Exception.Message) }
            } while (-not $finished)
            $reply = @($worker.EndInvoke($handle))
            if ($worker.HadErrors -or $reply.Count -ne 1) {
                throw "HTTP worker failed: $($worker.Streams.Error | Out-String)"
            }
            [pscustomobject]@{ Reply = $reply[0]; Samples = $samples.ToArray(); SampleErrors = $sampleErrors.ToArray() }
        }
        finally { if (-not $handle.IsCompleted) { $worker.Stop() } }
    }
    finally { $worker.Dispose() }
}

# Verify the API and resource counters before beginning a potentially long run.
try {
    $version = Invoke-ControlRequest '/api/version' $null
    $installed = Invoke-ControlRequest '/api/tags' $null
    $running = Invoke-ControlRequest '/api/ps' $null
    $baseline = Get-SystemSnapshot
}
catch { throw "Preflight failed. Start Ollama and check Windows CIM counters. $($_.Exception.Message)" }
if (@($running.models).Count -gt 0) {
    throw 'Stop currently loaded models with ollama stop <model>, then rerun. The benchmark requires an idle Ollama server.'
}

$runId = (Get-Date -Format 'yyyyMMdd_HHmmss_fff') + '_' + ([guid]::NewGuid().ToString('N').Substring(0, 6))
$runDirectory = Join-Path $OutputDirectory $runId
$null = New-Item -ItemType Directory -Path $runDirectory -Force
$runDirectory = (Resolve-Path -LiteralPath $runDirectory).Path
$csvPath = Join-Path $runDirectory 'measurements.csv'
$samplesPath = Join-Path $runDirectory 'resource-samples.csv'
Write-JsonFile ([ordered]@{
    StartedUtc = [DateTime]::UtcNow.ToString('o'); Models = $Models; Prompts = $Prompts
    OllamaVersion = $version; InstalledModels = $installed; PowerShellVersion = $PSVersionTable.PSVersion.ToString()
    Computer = @(Get-CimInstance Win32_ComputerSystem | Select-Object -Property @('Manufacturer', 'Model', 'TotalPhysicalMemory'))
    Processor = @(Get-CimInstance Win32_Processor | Select-Object -Property @('Name', 'NumberOfCores', 'NumberOfLogicalProcessors'))
    OS = @(Get-CimInstance Win32_OperatingSystem | Select-Object -Property @('Caption', 'Version', 'BuildNumber'))
    Baseline = $baseline; ContextWindow = $ContextWindow; NumPredict = $NumPredict
    Temperature = 0; Seed = 42; ThinkingMode = 'ModelDefault'; SampleIntervalMs = $SampleIntervalMs
    RequestTimeoutSeconds = $RequestTimeoutSeconds; ApiUrl = $BaseUrl
    ResourceScope = 'Total system; sampled through the HTTP request, including prompt processing and generation'
    TimingScope = 'Client HTTP round trip; excludes explicit model preload and result file writes'
}) (Join-Path $runDirectory 'run-info.json')

foreach ($model in $Models) {
    Write-Host "`n=== $model ===" -ForegroundColor Cyan
    $safeModel = $model -replace '[^A-Za-z0-9._-]', '_'
    $modelDirectory = Join-Path $runDirectory $safeModel
    $null = New-Item -ItemType Directory -Path $modelDirectory -Force
    $modelError = $null
    $preloadSeconds = $null
    $mayBeLoaded = $false
    try {
        try {
            $match = @($installed.models | Where-Object { $_.name -eq $model -or $_.model -eq $model })
            if ($match.Count -eq 0) { throw "Model is not installed. Run: ollama pull $model" }
            $show = Invoke-ControlRequest '/api/show' @{ model = $model }
            Write-JsonFile $show (Join-Path $modelDirectory 'model-info.json')
            $mayBeLoaded = $true
            $loadTimer = [Diagnostics.Stopwatch]::StartNew()
            $load = Invoke-ControlRequest '/api/generate' @{
                model = $model; prompt = ''; stream = $false; keep_alive = -1
                options = @{ num_ctx = $ContextWindow }
            }
            $loadTimer.Stop()
            if ($load.error) { throw [string]$load.error }
            $preloadSeconds = $loadTimer.Elapsed.TotalSeconds
            Write-JsonFile $load (Join-Path $modelDirectory 'preload.json')
            Write-JsonFile (Invoke-ControlRequest '/api/ps' $null) (Join-Path $modelDirectory 'loaded-models.json')
            Write-Host ('Loaded in {0:N2} seconds.' -f $preloadSeconds)
        }
        catch {
            $modelError = $_.Exception.Message
            Write-Warning "$model : $modelError"
        }

        foreach ($prompt in $Prompts) {
            Write-Host "  $($prompt.Id): $($prompt.Title)"
            $prefix = Join-Path $modelDirectory $prompt.Id
            $row = [ordered]@{
                RunId = $runId; Model = $model; PromptId = $prompt.Id; PromptTitle = $prompt.Title
                Status = 'Skipped'; Error = $modelError; StartedUtc = $null
                GenerationTokensPerSecond = $null; EndToEndSeconds = $null
                AverageCpuPercent = $null; PeakCpuPercent = $null
                AverageSystemMemoryPercent = $null; PeakSystemMemoryPercent = $null
                ResourceSampleCount = 0; SamplingErrors = $null; PreloadSeconds = $preloadSeconds
                OutputTokens = $null; InputTokens = $null; GenerationSeconds = $null
                OllamaTotalSeconds = $null; RequestLoadSeconds = $null; DoneReason = $null
                ContextWindow = $ContextWindow; NumPredict = $NumPredict; ThinkingMode = 'ModelDefault'
                ResponseFile = $null; ApiResponseFile = $null
            }
            if ($null -eq $modelError) {
                $body = @{
                    model = $model; prompt = $prompt.Text; stream = $false; keep_alive = -1
                    options = @{ temperature = 0; seed = 42; num_ctx = $ContextWindow; num_predict = $NumPredict }
                }
                Write-JsonFile $body "$prefix.request.json"
                $measured = $null
                try {
                    $measured = Invoke-MeasuredPrompt $body $model $prompt.Id
                    $reply = $measured.Reply
                    $row.StartedUtc = $reply.StartedUtc
                    $row.EndToEndSeconds = [Math]::Round($reply.EndToEndSeconds, 6)
                    $row.ResourceSampleCount = @($measured.Samples).Count
                    $row.SamplingErrors = $measured.SampleErrors -join ' | '
                    if ($row.ResourceSampleCount -gt 0) {
                        $measured.Samples | Export-Csv -LiteralPath $samplesPath -NoTypeInformation -Append -Encoding UTF8
                        $cpuIntegral = 0.0; $intervalTotal = 0.0
                        foreach ($sample in $measured.Samples) {
                            $cpuIntegral += $sample.CpuPercent * $sample.IntervalSeconds
                            $intervalTotal += $sample.IntervalSeconds
                        }
                        $row.AverageCpuPercent = [Math]::Round($cpuIntegral / $intervalTotal, 3)
                        $row.PeakCpuPercent = [Math]::Round(($measured.Samples | Measure-Object CpuPercent -Maximum).Maximum, 3)
                        $row.AverageSystemMemoryPercent = [Math]::Round(($measured.Samples | Measure-Object SystemMemoryPercent -Average).Average, 3)
                        $row.PeakSystemMemoryPercent = [Math]::Round(($measured.Samples | Measure-Object SystemMemoryPercent -Maximum).Maximum, 3)
                    }
                    if ($null -ne $reply.RawBody) {
                        [IO.File]::WriteAllText("$prefix.api-response.json", $reply.RawBody, $Utf8)
                        $row.ApiResponseFile = "$prefix.api-response.json"
                    }
                    if ($reply.ErrorText) { throw $reply.ErrorText }
                    $result = $reply.RawBody | ConvertFrom-Json
                    if ($result.error) { throw [string]$result.error }
                    [IO.File]::WriteAllText("$prefix.response.txt", [string]$result.response, $Utf8)
                    $row.ResponseFile = "$prefix.response.txt"
                    if ($result.thinking) { [IO.File]::WriteAllText("$prefix.thinking.txt", [string]$result.thinking, $Utf8) }
                    $row.OutputTokens = $result.eval_count; $row.InputTokens = $result.prompt_eval_count
                    $row.DoneReason = $result.done_reason
                    if ($null -ne $result.total_duration) { $row.OllamaTotalSeconds = [double]$result.total_duration / 1e9 }
                    if ($null -ne $result.load_duration) { $row.RequestLoadSeconds = [double]$result.load_duration / 1e9 }
                    $row.Status = 'Completed'; $row.Error = $null
                    if ([double]$result.eval_duration -gt 0 -and $null -ne $result.eval_count) {
                        $row.GenerationSeconds = [double]$result.eval_duration / 1e9
                        $row.GenerationTokensPerSecond = [Math]::Round([double]$result.eval_count / $row.GenerationSeconds, 6)
                    }
                    else { $row.Status = 'MissingMetrics'; $row.Error = 'Ollama did not return valid generation timing/token counts.' }
                    if (-not $result.done) { $row.Status = 'Incomplete'; $row.Error = 'API did not confirm completed generation.' }
                    elseif ($result.done_reason -eq 'length') { $row.Status = 'Truncated'; $row.Error = 'Output reached a length limit; review or rerun with a higher NumPredict.' }
                    elseif ([string]::IsNullOrWhiteSpace([string]$result.response)) { $row.Status = 'EmptyResponse'; $row.Error = 'No answer text returned. Inspect thinking/API response.' }
                    if ($row.ResourceSampleCount -eq 0 -or $row.SamplingErrors) {
                        $row.Error = @($row.Error, 'Resource sampling was incomplete; inspect SamplingErrors.') -join ' '
                        if ($row.Status -eq 'Completed') { $row.Status = 'SamplingWarning' }
                    }
                    Write-Host ('    {0} | {1} tokens/sec | {2:N2} sec | CPU avg {3}% | RAM peak {4}%' -f $row.Status, $row.GenerationTokensPerSecond, $row.EndToEndSeconds, $row.AverageCpuPercent, $row.PeakSystemMemoryPercent)
                }
                catch {
                    $row.Status = 'Failed'; $row.Error = $_.Exception.Message
                    [IO.File]::WriteAllText("$prefix.error.txt", $row.Error, $Utf8)
                    Write-Warning "$model / $($prompt.Id): $($row.Error)"
                    # A timed-out server may still be cancelling generation. Skip
                    # this model's remaining requests and unload before advancing.
                    $modelError = "Skipped after request failure: $($row.Error)"
                }
            }
            [pscustomobject]$row | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Append -Encoding UTF8
        }
    }
    finally {
        if ($mayBeLoaded) { Stop-BenchmarkModel $model }
    }
}

Write-Host "`nFinished. Measurements: $csvPath" -ForegroundColor Green
Write-Host "Responses and supporting files: $runDirectory"
