param(
    [UInt64] $Mask = 65535,
    [int] $Iterations = 30,
    [int] $Warmup = 12,
    [int] $Exports = 16,
    [string[]] $Scenarios = @('Csv25K','Excel25K','CsvJsonAlways','CsvQuotesAlways'),
    [string] $ScenarioList,
    [string] $Label = 'qualification',
    [switch] $Plan
)
$ErrorActionPreference = 'Stop'
if ($ScenarioList) { $Scenarios = $ScenarioList.Split(',') }
Import-Module PSPublishModule
$env:OFFICEIMO_PERFORMANCE_SNAPSHOTS = $PSScriptRoot
$env:OFFICEIMO_BENCHMARK_DATA = Join-Path $PSScriptRoot 'fixtures'
$env:OFFICEIMO_BENCHMARK_OUTPUT = Join-Path $PSScriptRoot 'file-output'
[Reflection.Assembly]::LoadFrom((Join-Path $PSScriptRoot 'compare/bin/Release/net10.0/Compare.dll')) | Out-Null
[Diagnostics.Process]::GetCurrentProcess().ProcessorAffinity = [IntPtr]::new([long]$Mask)
[Diagnostics.Process]::GetCurrentProcess().PriorityClass = 'Normal'
$global:OfficePerfProbes = @{}
$global:OfficePerfLengths = @{}
$global:OfficePerfExports = $Exports
foreach ($scenario in $Scenarios) {
    $probe = [VersionComparison]::new()
    $probe.Scenario = $scenario
    $probe.Setup()
    $global:OfficePerfProbes[$scenario] = $probe
    $global:OfficePerfLengths[$scenario] = $probe.Before()
}
try {
Invoke-BenchmarkSuite -OutputRoot (Join-Path $PSScriptRoot ($Label + '-' + $Mask)) -WarmupCount $Warmup -IterationCount $Iterations -RunOrder Rotated -OutlierMode None -Plan:$Plan -Settings {
    New-BenchmarkSuite 'Excel CSV milestone qualification' {
        Add-BenchmarkMetadata AffinityMask $Mask
        Add-BenchmarkMetadata ExportsPerOperation $Exports
        Add-BenchmarkMetadata ProcessId $PID
        Add-BenchmarkMetadata Priority ([Diagnostics.Process]::GetCurrentProcess().PriorityClass.ToString())
        if ($env:OFFICEIMO_PERFORMANCE_AA -eq '1') {
            Add-BenchmarkMetadata ComparisonMode 'Identical baseline assemblies'
        } else {
            Add-BenchmarkMetadata ComparisonMode 'Baseline versus candidate'
        }
        Add-BenchmarkCaseSource @($global:OfficePerfProbes.Keys | Sort-Object | ForEach-Object { [pscustomobject]@{Name=$_} })
        Set-BenchmarkSetup {
            param($case, $run)
            $run.Probe = $global:OfficePerfProbes[$case.Scenario]
            $run.Expected = $global:OfficePerfLengths[$case.Scenario]
        }
        Add-BenchmarkEngine Before {
            Add-BenchmarkOperation ExportBatch {
                param($case, $run)
                for($index=0; $index -lt $global:OfficePerfExports; $index++) { $run.Length = $run.Probe.Before() }
            }
        }
        Add-BenchmarkEngine After {
            Add-BenchmarkOperation ExportBatch {
                param($case, $run)
                for($index=0; $index -lt $global:OfficePerfExports; $index++) { $run.Length = $run.Probe.After() }
            }
        }
        Add-BenchmarkValidation {
            param($case, $run)
            if($run.Length -ne $run.Expected){ throw 'Output length differs from the validated baseline.' }
        }
        Add-BenchmarkComparison -Baseline Before -Metric MeanMs,MedianMs -TieTolerance 0.05
    }
}
} finally {
    foreach ($probe in $global:OfficePerfProbes.Values) { $probe.Dispose() }
}
