param(
    [string] $BinaryRoot = (Join-Path $PSScriptRoot '../../OfficeIMO.Drawing.Benchmarks/bin/Release/net10.0'),
    [string] $OutputRoot = (Join-Path $PSScriptRoot '../../.validation/release-quality-images'),
    [int] $WarmupCount = 1,
    [int] $IterationCount = 3,
    [string] $ReferenceSummaryPath,
    [double] $PerformanceTolerance = 0.35,
    [switch] $Plan
)

$ErrorActionPreference = 'Stop'
$BinaryRoot = (Resolve-Path -LiteralPath $BinaryRoot).Path
[void] [Reflection.Assembly]::LoadFrom((Join-Path $BinaryRoot 'OfficeIMO.Core.dll'))
Import-Module PSPublishModule -MinimumVersion 3.0.141 -ErrorAction Stop

$result = Invoke-BenchmarkSuite -Path (Join-Path $PSScriptRoot 'release-quality-images.benchmark.ps1') `
    -OutputRoot $OutputRoot -Variable @{ BinaryRoot = $BinaryRoot } `
    -WarmupCount $WarmupCount -IterationCount $IterationCount -Plan:$Plan
$result

if (-not $Plan -and @($result.Summary | Where-Object Status -ne 'Succeeded').Count -gt 0) {
    throw 'Release-quality image evidence failed validation. Inspect the benchmark artifacts.'
}

if (-not $Plan -and $ReferenceSummaryPath) {
    $ReferenceSummaryPath = (Resolve-Path -LiteralPath $ReferenceSummaryPath).Path
    $summaryPath = $result.Artifacts['summary.json']
    $gateRoot = Split-Path -Parent $summaryPath
    $gateMetrics = @(
        @{ Name = 'MedianMs'; Absolute = 10.0 },
        @{ Name = 'ManagedAllocatedBytes'; Absolute = 1048576.0 },
        @{ Name = 'PeakWorkingSetBytes'; Absolute = 16777216.0 },
        @{ Name = 'PeakPrivateBytes'; Absolute = 16777216.0 },
        @{ Name = 'PeakNativeBytesEstimate'; Absolute = 16777216.0 },
        @{ Name = 'EncodedBytes'; Absolute = 0.0 },
        @{ Name = 'CancellationLatencyMilliseconds'; Absolute = 5.0 },
        @{ Name = 'MeanAbsoluteError'; Absolute = 0.5 }
    )
    foreach ($metric in $gateMetrics) {
        $baselinePath = Join-Path $gateRoot ("reference-$($metric.Name).json")
        Test-BenchmarkGate -SummaryPath $ReferenceSummaryPath -BaselinePath $baselinePath `
            -Metric $metric.Name -GroupBy scenario -Update | Out-Null
        $gate = Test-BenchmarkGate -SummaryPath $summaryPath -BaselinePath $baselinePath `
            -Metric $metric.Name -GroupBy scenario -MetricDirection LowerIsBetter `
            -RelativeTolerance $PerformanceTolerance -AbsoluteToleranceMs $metric.Absolute -AllowNew
        if (-not $gate.Passed) {
            throw "Release-quality image regression gate failed for $($metric.Name)."
        }
    }
}
