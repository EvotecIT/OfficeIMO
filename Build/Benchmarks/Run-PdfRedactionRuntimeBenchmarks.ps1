param(
    [string] $BinaryRoot = (Join-Path $PSScriptRoot '../../OfficeIMO.Pdf.Benchmarks/bin/Release/net8.0'),
    [string] $OutputRoot = (Join-Path $PSScriptRoot '../../Ignore/Benchmarks/PdfRedactionRuntime'),
    [string] $ModulePath = 'PSPublishModule',
    [ValidateRange(1, 500)] [int[]] $Pages = @(1, 25, 100),
    [string] $InputPath = '',
    [string] $Pattern = 'private account [0-9]{3}',
    [ValidateRange(0, 100)] [int] $WarmupCount = 2,
    [ValidateRange(1, 100)] [int] $IterationCount = 5,
    [ValidateRange(0, 1000)] [int] $MemorySamplingIntervalMilliseconds = 0,
    [switch] $Plan
)
$ErrorActionPreference = 'Stop'
$BinaryRoot = (Resolve-Path -LiteralPath $BinaryRoot).Path
[void] [Reflection.Assembly]::LoadFrom((Join-Path $BinaryRoot 'OfficeIMO.Core.dll'))
[void] [Reflection.Assembly]::LoadFrom((Join-Path $BinaryRoot 'OfficeIMO.Pdf.dll'))
Import-Module $ModulePath -ErrorAction Stop
$result = Invoke-BenchmarkSuite -Path (Join-Path $PSScriptRoot 'pdf-redaction-runtime.benchmark.ps1') `
    -OutputRoot $OutputRoot -Variable @{ BinaryRoot = $BinaryRoot; Pages = $Pages; InputPath = $InputPath; Pattern = $Pattern } `
    -WarmupCount $WarmupCount -IterationCount $IterationCount `
    -MemorySamplingIntervalMilliseconds $MemorySamplingIntervalMilliseconds `
    -RunMode $(if ($MemorySamplingIntervalMilliseconds -gt 0) { "memory-$MemorySamplingIntervalMilliseconds-ms" } else { 'standard' }) -Plan:$Plan
$result
if (-not $Plan -and @($result.Samples | Where-Object Status -ne 'Succeeded').Count -gt 0) {
    throw 'Redaction runtime validation failed. Inspect the retained artifacts.'
}
if (-not $Plan -and @($result.Samples | Where-Object { $null -eq $_.AllocatedBytes }).Count -gt 0) {
    throw 'Managed allocation measurement is missing from the shared runner.'
}
if (-not $Plan -and $MemorySamplingIntervalMilliseconds -gt 0 -and @($result.Samples | Where-Object {
    -not $_.Metrics.ContainsKey('MemorySampleCount') -or $_.Metrics['MemorySampleCount'] -lt 2 -or
    $_.Metrics['MemorySamplingFailed'] -ne 0
}).Count -gt 0) {
    throw 'Requested operation memory observations are missing or invalid.'
}
