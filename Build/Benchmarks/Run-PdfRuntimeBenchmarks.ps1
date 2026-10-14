param(
    [string] $BinaryRoot = (Join-Path $PSScriptRoot '../../OfficeIMO.Pdf.Benchmarks/bin/Release/net10.0'),
    [string] $OutputRoot = (Join-Path $PSScriptRoot '../../Ignore/Benchmarks/PdfRuntime'),
    [ValidateRange(1, 500)] [int[]] $Pages = @(1, 20, 100),
    [string[]] $Operation,
    [ValidateRange(0, 100)] [int] $WarmupCount = 2,
    [ValidateRange(1, 100)] [int] $IterationCount = 7,
    [ValidateRange(0, 1000)] [int] $MemorySamplingIntervalMilliseconds = 0,
    [Nullable[UInt64]] $ProcessorAffinityMask,
    [ValidateSet('Idle', 'BelowNormal', 'Normal', 'AboveNormal', 'High')] [string] $ProcessPriority = 'Normal',
    [string] $TextTool = 'pdftotext',
    [string] $EnvironmentEvidence = '',
    [switch] $Plan
)
$ErrorActionPreference = 'Stop'
$PSNativeCommandUseErrorActionPreference = $true
$BinaryRoot = (Resolve-Path -LiteralPath $BinaryRoot).Path
$OutputRoot = [IO.Path]::GetFullPath($OutputRoot)
if (@($Pages | Sort-Object -Unique).Count -ne $Pages.Count) { throw 'Page counts must be distinct.' }
if ($EnvironmentEvidence) { $EnvironmentEvidence = (Resolve-Path -LiteralPath $EnvironmentEvidence).Path }
$textCommand = Get-Command $TextTool -CommandType Application -ErrorAction Stop
$TextTool = $textCommand.Source
$textVersion = (& $TextTool -v 2>&1 | Out-String).Trim()
if ($LASTEXITCODE -ne 0) { throw 'Cannot identify the independent text reader.' }

# A fresh PowerShell process avoids importing a different OfficeIMO.Core through the runner.
foreach ($name in 'OfficeIMO.Core', 'OfficeIMO.Pdf') {
    [void] [Reflection.Assembly]::LoadFrom((Join-Path $BinaryRoot "$name.dll"))
}
Import-Module PSPublishModule -MinimumVersion 3.0.168 -ErrorAction Stop
$parameters = @{
    Path = Join-Path $PSScriptRoot 'pdf-runtime.benchmark.ps1'
    OutputRoot = $OutputRoot
    Variable = @{
        BinaryRoot = $BinaryRoot; Pages = $Pages; ValidationRoot = Join-Path $OutputRoot 'inputs'
        TextTool = $TextTool; TextToolVersion = $textVersion; EnvironmentEvidence = $EnvironmentEvidence
    }
    WarmupCount = $WarmupCount
    IterationCount = $IterationCount
    RunOrder = 'Rotated'
    OutlierMode = 'None'
    MemorySamplingIntervalMilliseconds = $MemorySamplingIntervalMilliseconds
    RunMode = $(if ($MemorySamplingIntervalMilliseconds -gt 0) { "memory-$MemorySamplingIntervalMilliseconds-ms" } else { 'standard' })
    Plan = $Plan
}
if ($Operation) { $parameters.Operation = $Operation }
if ($null -ne $ProcessorAffinityMask) { $parameters.ProcessorAffinityMask = [UInt64] $ProcessorAffinityMask }
if ($IsWindows) { $parameters.ProcessPriority = $ProcessPriority }
$result = Invoke-BenchmarkSuite @parameters
$result
if (-not $Plan -and @($result.Samples | Where-Object Status -ne 'Succeeded').Count -gt 0) {
    throw 'PDF runtime validation failed. Inspect the raw samples and retained artifacts.'
}
if (-not $Plan -and @($result.Samples | Where-Object { $null -eq $_.AllocatedBytes }).Count -gt 0) {
    throw 'The shared runner did not record managed allocation.'
}
if (-not $Plan -and $MemorySamplingIntervalMilliseconds -gt 0 -and @($result.Samples | Where-Object {
    -not $_.Metrics.ContainsKey('MemorySampleCount') -or $_.Metrics['MemorySampleCount'] -lt 2 -or
    $_.Metrics['MemorySamplingFailed'] -ne 0
}).Count -gt 0) {
    throw 'Requested operation memory observations are missing or invalid.'
}
