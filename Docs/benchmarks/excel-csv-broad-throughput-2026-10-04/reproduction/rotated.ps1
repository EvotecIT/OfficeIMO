param([long] $Mask = 65535, [int] $Iterations = 24, [int] $Warmup = 12, [int] $Operations = 4,
    [string] $Cases = 'rotated-cases.json', [string] $Label = 'rotated', [switch] $IdenticalBaseline)
$ErrorActionPreference = 'Stop'
Import-Module PSPublishModule
$env:OFFICEIMO_PERFORMANCE_SNAPSHOTS = $PSScriptRoot
$env:OFFICEIMO_COMPARISON_CASES = Join-Path $PSScriptRoot $Cases
$env:OFFICEIMO_PERFORMANCE_AA = if ($IdenticalBaseline) { '1' } else { '0' }
$env:OFFICEIMO_BENCHMARK_DATA = Join-Path $PSScriptRoot 'fixtures'
$env:OFFICEIMO_BENCHMARK_OUTPUT = Join-Path $PSScriptRoot 'files'
[Diagnostics.Process]::GetCurrentProcess().ProcessorAffinity = [IntPtr]::new($Mask)
[Diagnostics.Process]::GetCurrentProcess().PriorityClass = 'Normal'
[Reflection.Assembly]::LoadFrom((Join-Path $PSScriptRoot 'compare/bin/Release/net10.0/Compare.dll')) | Out-Null
$global:BroadProbes = @{}
$global:BroadExpected = @{}
$global:BroadOperations = $Operations
foreach ($key in [SnapshotProbe]::Keys()) {
    $before = [SnapshotProbe]::Load($key, 'baseline')
    $after = [SnapshotProbe]::Load($key, 'candidate')
    $expected = $before.RunSynchronously().ToString()
    if ($after.RunSynchronously().ToString() -ne $expected) { throw "Validation failed: $key" }
    $global:BroadProbes[$key] = @{ Before = $before; After = $after }
    $global:BroadExpected[$key] = $expected
}
try {
    $result = Invoke-BenchmarkSuite -OutputRoot (Join-Path $PSScriptRoot "$Label-$Mask") -WarmupCount $Warmup -IterationCount $Iterations -RunOrder Rotated -OutlierMode None -Settings {
        New-BenchmarkSuite 'Excel CSV broad qualification' {
            Add-BenchmarkMetadata AffinityMask $Mask
            Add-BenchmarkMetadata OperationsPerSample $Operations
            Add-BenchmarkMetadata IdenticalBaseline ([string]$IdenticalBaseline)
            Add-BenchmarkMetadata Priority ([Diagnostics.Process]::GetCurrentProcess().PriorityClass.ToString())
            Add-BenchmarkCaseSource @($global:BroadProbes.Keys | Sort-Object | ForEach-Object { [pscustomobject]@{Name=$_} })
            Set-BenchmarkSetup {
                param($case, $run)
                $run.Probe = $global:BroadProbes[$case.Scenario][$case.Engine]
                $run.Expected = $global:BroadExpected[$case.Scenario]
            }
            foreach ($engine in 'Before','After') {
                Add-BenchmarkEngine $engine {
                    Add-BenchmarkOperation ReadBatch {
                        param($case, $run)
                        for ($i=0; $i -lt $global:BroadOperations; $i++) { $run.Result = $run.Probe.RunSynchronously() }
                    }
                }
            }
            Add-BenchmarkValidation {
                param($case, $run)
                if ($run.Result.ToString() -ne $run.Expected) { throw 'Output differs from validated setup.' }
            }
            Add-BenchmarkComparison -Baseline Before -Metric MeanMs,MedianMs -TieTolerance 0.05
        }
    }
    if (@($result.Samples | Where-Object Status -ne 'Succeeded').Count -gt 0) {
        throw 'Comparison contains failed samples; inspect the retained artifacts.'
    }
    $result
} finally {
    foreach ($pair in $global:BroadProbes.Values) { $pair.Before.Dispose(); $pair.After.Dispose() }
}
