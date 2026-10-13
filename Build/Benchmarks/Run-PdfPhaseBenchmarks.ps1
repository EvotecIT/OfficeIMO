param(
    [Parameter(Mandatory)] [ValidateSet('RenderStages', 'FirstWrite', 'ClosedLedger')] [string] $Family,
    [string] $BinaryRoot = (Join-Path $PSScriptRoot '../../OfficeIMO.Pdf.Benchmarks/bin/Release/net10.0'),
    [string] $OutputRoot = (Join-Path $PSScriptRoot '../../Ignore/Benchmarks/PdfPhases'),
    [ValidateRange(1, 500)] [int[]] $Pages,
    [ValidateSet(0, 90)] [int[]] $Rotation = @(0, 90),
    [ValidateSet('DetectedProportions', 'ClosedRuling')] [string] $ExpectedGeometry,
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
if (-not $Pages) { $Pages = if ($Family -eq 'FirstWrite') { @(20, 100, 500) } else { @(20) } }
if (@($Pages | Sort-Object -Unique).Count -ne $Pages.Count) { throw 'Page counts must be distinct.' }
if ($Family -eq 'RenderStages' -and ($Pages.Count -ne 1 -or $Pages[0] -ne 20)) { throw 'Render stages use the fixed 20-page, four-preview corpus.' }
if ($Family -eq 'ClosedLedger' -and -not $ExpectedGeometry) { throw 'ClosedLedger requires an explicit expected geometry contract.' }
if (@($Rotation | Sort-Object -Unique).Count -ne $Rotation.Count) { throw 'Rotations must be distinct.' }
$BinaryRoot = (Resolve-Path -LiteralPath $BinaryRoot).Path
$OutputRoot = [IO.Path]::GetFullPath((Join-Path $OutputRoot $Family))
if ($EnvironmentEvidence) { $EnvironmentEvidence = (Resolve-Path -LiteralPath $EnvironmentEvidence).Path }
$TextTool = (Get-Command $TextTool -CommandType Application -ErrorAction Stop).Source
$textVersion = (& $TextTool -v 2>&1 | Out-String).Trim()
if ($LASTEXITCODE -ne 0) { throw 'Cannot identify the independent text reader.' }
foreach ($name in 'OfficeIMO.Core', 'OfficeIMO.Pdf') {
    [void] [Reflection.Assembly]::LoadFrom((Join-Path $BinaryRoot "$name.dll"))
}
Import-Module PSPublishModule -MinimumVersion 3.0.168 -ErrorAction Stop
$parameters = @{
    Path = Join-Path $PSScriptRoot 'pdf-phases.benchmark.ps1'
    OutputRoot = $OutputRoot
    Variable = @{
        Family = $Family; BinaryRoot = $BinaryRoot; Pages = $Pages; Rotation = $Rotation; ExpectedGeometry = $ExpectedGeometry
        ValidationRoot = Join-Path $OutputRoot 'inputs'; TextTool = $TextTool
        TextToolVersion = $textVersion; EnvironmentEvidence = $EnvironmentEvidence
    }
    WarmupCount = $WarmupCount; IterationCount = $IterationCount; RunOrder = 'Rotated'; OutlierMode = 'None'
    MemorySamplingIntervalMilliseconds = $MemorySamplingIntervalMilliseconds
    RunMode = $(if ($MemorySamplingIntervalMilliseconds -gt 0) { "diagnostic-memory-$MemorySamplingIntervalMilliseconds-ms" } else { 'diagnostic' })
    Plan = $Plan
}
if ($Operation) { $parameters.Operation = $Operation }
if ($null -ne $ProcessorAffinityMask) { $parameters.ProcessorAffinityMask = [UInt64] $ProcessorAffinityMask }
if ($IsWindows) { $parameters.ProcessPriority = $ProcessPriority }
$result = Invoke-BenchmarkSuite @parameters
$result
if (-not $Plan -and @($result.Samples | Where-Object Status -ne 'Succeeded').Count -gt 0) {
    throw 'PDF phase evidence validation failed. Inspect the raw samples and retained artifacts.'
}
if (-not $Plan -and @($result.Samples | Where-Object { $null -eq $_.AllocatedBytes }).Count -gt 0) {
    throw 'The shared runner did not record managed allocation.'
}
if (-not $Plan -and $MemorySamplingIntervalMilliseconds -gt 0 -and @($result.Samples | Where-Object {
    -not $_.Metrics.ContainsKey('MemorySampleCount') -or $_.Metrics['MemorySampleCount'] -lt 2 -or $_.Metrics['MemorySamplingFailed'] -ne 0
}).Count -gt 0) { throw 'Requested operation memory observations are missing or invalid.' }
