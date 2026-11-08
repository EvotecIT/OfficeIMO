param(
    [string] $BinaryRoot = (Join-Path $PSScriptRoot '../../OfficeIMO.Migration.Benchmarks/bin/Release/net10.0'),
    [string] $OutputRoot = (Join-Path $PSScriptRoot '../../Ignore/Benchmarks/Migration'),
    [string] $ModulePath = 'PSPublishModule',
    [string] $ProductSource = '',
    [ValidateSet('Word69', 'Word69Notes', 'Word95Notes', 'Biff4', 'IWork09', 'Dbx', 'OfficeXml')] [string[]] $BenchmarkProfile = @('Word69', 'Biff4', 'IWork09', 'Dbx', 'OfficeXml'),
    [ValidateSet('Small', 'Medium', 'Large')] [string[]] $Scale = @('Small', 'Medium', 'Large'),
    [ValidateSet('LoadInspect', 'ConvertSave', 'SearchContent')] [string[]] $Operation = @('LoadInspect', 'ConvertSave'),
    [ValidateRange(0, 100)] [int] $WarmupCount = 2,
    [ValidateRange(1, 100)] [int] $IterationCount = 5,
    [ulong] $ProcessorAffinityMask = 0,
    [ValidateRange(0, 1000)] [int] $MemorySamplingIntervalMilliseconds = 0,
    [switch] $MeasureRetainedMemory,
    [switch] $Plan
)
$ErrorActionPreference = 'Stop'
$BinaryRoot = (Resolve-Path -LiteralPath $BinaryRoot).Path
foreach ($name in 'OfficeIMO.Core', 'OfficeIMO.Rtf', 'OfficeIMO.Word', 'OfficeIMO.Excel', 'OfficeIMO.Email', 'OfficeIMO.IWork', 'OfficeIMO.Word.IWork') {
    $path = Join-Path $BinaryRoot "$name.dll"
    $assembly = [Reflection.Assembly]::LoadFrom($path)
    if ($assembly.Location -ne $path) { throw "The isolated benchmark process bound $name from another directory." }
}
Import-Module $ModulePath -ErrorAction Stop
$caseNames = @(foreach ($family in $BenchmarkProfile) { foreach ($size in $Scale) { "$family-$size" } })
$settings = @{
    Path = Join-Path $PSScriptRoot 'migration.benchmark.ps1'
    OutputRoot = $OutputRoot
    Variable = @{ BinaryRoot = $BinaryRoot; Profiles = $BenchmarkProfile; Scales = $Scale; MeasureRetainedMemory = [bool] $MeasureRetainedMemory; ProductSource = $ProductSource }
    Case = $caseNames
    Operation = $Operation
    WarmupCount = $WarmupCount
    IterationCount = $IterationCount
    MemorySamplingIntervalMilliseconds = $MemorySamplingIntervalMilliseconds
    RunMode = $(if ($MemorySamplingIntervalMilliseconds -gt 0) { "memory-$MemorySamplingIntervalMilliseconds-ms" } else { 'standard' })
    Plan = $Plan
}
if ($ProcessorAffinityMask -ne 0) { $settings.ProcessorAffinityMask = $ProcessorAffinityMask }
$result = Invoke-BenchmarkSuite @settings
$result
if (-not $Plan -and @($result.Samples | Where-Object Status -eq 'Failed').Count -gt 0) {
    throw 'Migration workload validation failed. Inspect the retained benchmark artifacts.'
}
if (-not $Plan -and @($result.Samples | Where-Object { $_.Status -eq 'Succeeded' -and $null -eq $_.AllocatedBytes }).Count -gt 0) {
    throw 'The shared runner did not provide managed allocation measurements.'
}


if (-not $Plan -and $MeasureRetainedMemory -and @($result.Samples | Where-Object {
    $_.Status -eq 'Succeeded' -and -not $_.Metrics.ContainsKey('RetainedManagedDeltaBytes')
}).Count -gt 0) {
    throw 'The requested collected managed-heap observation is missing.'
}
if (-not $Plan -and $MemorySamplingIntervalMilliseconds -gt 0 -and @($result.Samples | Where-Object {
    $_.Status -eq 'Succeeded' -and (-not $_.Metrics.ContainsKey('MemorySampleCount') -or $_.Metrics['MemorySampleCount'] -lt 2 -or
        $_.Metrics['MemorySamplingFailed'] -ne 0 -or $_.Metrics['MemorySamplingIntervalMs'] -ne $MemorySamplingIntervalMilliseconds)
}).Count -gt 0) {
    throw 'The requested operation memory observations are missing or invalid.'
}
