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
    $referenceRows = @(Get-Content -LiteralPath $ReferenceSummaryPath -Raw | ConvertFrom-Json)
    $currentRows = @(Get-Content -LiteralPath $summaryPath -Raw | ConvertFrom-Json)

    function Get-RowMetricValue {
        param($Row, [string] $Name)
        if ($Row.PSObject.Properties.Name -contains $Name) { return $Row.$Name }
        if ($null -ne $Row.metrics -and $Row.metrics.PSObject.Properties.Name -contains $Name) { return $Row.metrics.$Name }
        return $null
    }

    $inputHashMetrics = @(0..7 | ForEach-Object { "InputHashWord$_" })
    $referenceInputRows = @($referenceRows | Where-Object {
        $_.variables.Workload -ne 'Encode' -and $_.variables.Workload -ne 'Resample'
    })
    $referenceProvenanceRows = @($referenceInputRows | Where-Object {
        $row = $_
        @($inputHashMetrics | Where-Object { $null -ne (Get-RowMetricValue $row $_) }).Count -eq $inputHashMetrics.Count
    })
    if ($referenceInputRows.Count -ne $referenceProvenanceRows.Count) {
        Write-Host 'Skipping regression comparison because the reference summary predates input-provenance metrics.'
        return
    }

    foreach ($current in $currentRows) {
        $workload = $current.variables.Workload
        if ($workload -eq 'Encode' -or $workload -eq 'Resample') { continue }
        $reference = $referenceRows | Where-Object scenario -eq $current.scenario | Select-Object -First 1
        if ($null -eq $reference) { continue }
        foreach ($hashMetric in $inputHashMetrics) {
            $referenceWord = Get-RowMetricValue $reference $hashMetric
            $currentWord = Get-RowMetricValue $current $hashMetric
            if ($null -eq $currentWord -or $currentWord -ne $referenceWord) {
                throw "Release-quality image input provenance changed for $($current.scenario); performance comparison is not valid."
            }
        }
    }
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
        if (-not ($referenceRows | Where-Object { $null -ne (Get-RowMetricValue $_ $metric.Name) } | Select-Object -First 1)) {
            Write-Host "Skipping $($metric.Name) regression comparison because the reference summary predates that metric."
            continue
        }
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
