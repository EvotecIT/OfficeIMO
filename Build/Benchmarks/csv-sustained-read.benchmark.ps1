$repositoryRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$binaryRoot = Get-BenchmarkInput BinaryRoot (Join-Path $repositoryRoot 'OfficeIMO.CSV/bin/Release/net10.0')
$fixtureRoot = Get-BenchmarkInput FixtureRoot -Required
$rows = Get-BenchmarkInput Rows 100000, 1000000 -Int
$degree = Get-BenchmarkInput Degree 4 -Int
$batch = Get-BenchmarkInput Batch 1024 -Int
$sampleMemory = Get-BenchmarkInput SampleMemory $false -Bool
$selectedEngineSetup = Get-BenchmarkInput SelectedEngineSetup $false -Bool
$references = @([IO.Directory]::GetFiles((Join-Path $PSHOME 'ref'), '*.dll'))
foreach ($name in 'OfficeIMO.Core', 'OfficeIMO.CSV') {
    $path = Join-Path $binaryRoot "$name.dll"
    $loaded = [Reflection.Assembly]::LoadFrom($path)
    if ((Get-FileHash $loaded.Location -Algorithm SHA256).Hash -ne (Get-FileHash $path -Algorithm SHA256).Hash) {
        throw "The process already loaded a different $name binary. Run this lane in a fresh process."
    }
    $references += $path
}
$references += [PowerForge.BenchmarkMemoryProbe].Assembly.Location
Add-Type -Path (Join-Path $PSScriptRoot 'CsvSustainedReadWorkload.cs') -ReferencedAssemblies $references
$workloads = @{}

New-BenchmarkSuite 'officeimo-csv-sustained-read' {
    Set-BenchmarkPolicy -Warmup 2 -Iteration 7 -Order Rotated -OutlierMode None -MemoryCleanup BeforeIteration
    Add-BenchmarkMetadata Contract 'FirstRow/AllRowsAsync: async initialization/read, every string field consumed. Typed: async initialization, synchronous ordered projection, full checksum, no retained row array.'
    Add-BenchmarkMetadata SampleMemory ([string] $sampleMemory)
    Add-BenchmarkMetadata SelectedEngineSetup ([string] $selectedEngineSetup)
    Add-BenchmarkMetadata ProcessId ([string] [Diagnostics.Process]::GetCurrentProcess().Id)
    Add-BenchmarkMetadata Degree ([string] $degree[0])
    Add-BenchmarkMetadata Batch ([string] $batch[0])
    Add-BenchmarkMetadata Sampling 'PowerForge 5ms sampled lower-bound peaks; separate instrumented memory lane.'
    Add-BenchmarkMetadata BinarySha256 (Get-FileHash (Join-Path $binaryRoot 'OfficeIMO.CSV.dll') -Algorithm SHA256).Hash
    Add-BenchmarkMetadata CoreBinarySha256 (Get-FileHash (Join-Path $binaryRoot 'OfficeIMO.Core.dll') -Algorithm SHA256).Hash
    Add-BenchmarkMetadata PowerForgeBinarySha256 (Get-FileHash ([PowerForge.BenchmarkMemoryProbe].Assembly.Location) -Algorithm SHA256).Hash
    Add-BenchmarkMetadata PowerShellRunnerBinarySha256 (Get-FileHash ([PowerForge.PowerShellBenchmarkRunner].Assembly.Location) -Algorithm SHA256).Hash
    Add-BenchmarkMetadata WorkloadSha256 (Get-FileHash (Join-Path $PSScriptRoot 'CsvSustainedReadWorkload.cs') -Algorithm SHA256).Hash
    Add-BenchmarkMetadata AffinityMask $(if ([OperatingSystem]::IsWindows() -or [OperatingSystem]::IsLinux()) { [Diagnostics.Process]::GetCurrentProcess().ProcessorAffinity.ToInt64() } else { 'unqualified' })
    Add-BenchmarkMetadata Priority ([Diagnostics.Process]::GetCurrentProcess().PriorityClass.ToString())
    Add-BenchmarkCaseSource {
        foreach ($count in $rows) {
            foreach ($shape in 'Plain', 'Multiline') {
                [pscustomobject]@{ Name = "Rows-$count-$shape"; Rows = $count; Shape = $shape }
            }
        }
    }
    Set-BenchmarkSetup {
        param($case, $run)
        $key = "$($case.Rows)-$($case.Shape)"
        if (-not $workloads.ContainsKey($key)) {
            $workloads[$key] = [CsvSustainedReadWorkload]::new((Join-Path $fixtureRoot "$key.csv"), $case.Rows, ($case.Shape -eq 'Multiline'), $degree[0], $batch[0])
        }
        $run.Workload = $workloads[$key]
        $validationKey = "$key-$($case.Operation)-$(if ($selectedEngineSetup) { $case.Engine } else { 'Both' })"
        if (-not $workloads.ContainsKey($validationKey)) {
            if ($selectedEngineSetup) { $run.Workload.Prepare($case.Operation, $case.Engine) }
            else { $run.Workload.Prepare($case.Operation) }
            $workloads[$validationKey] = $true
        }
    }
    foreach ($engine in 'Snapshot', 'Incremental') {
        Add-BenchmarkEngine $engine {
            foreach ($operation in 'FirstRow', 'AllRowsAsync', 'TypedSequential', 'TypedParallel') {
                Add-BenchmarkOperation $operation {
                    param($case, $run)
                    $run.Workload.Execute(($case.Engine -eq 'Incremental'), $case.Operation, $sampleMemory)
                }
            }
        }
    }
    Add-BenchmarkValidation { param($case, $run) $run.Workload.Validate() }
    Add-BenchmarkMetric InputBytes { param($case, $run) $run.Workload.InputBytes }
    if ($sampleMemory) {
        foreach ($metric in 'BaselineManagedBytes', 'BaselineResidentBytes', 'PeakManagedBytes', 'PeakResidentBytes', 'ManagedIncreaseBytes', 'ResidentIncreaseBytes') {
            Add-BenchmarkMetric $metric ({ param($case, $run) $run.Workload.Memory.($metric) }.GetNewClosure())
        }
        Add-BenchmarkMetric MemorySampleCount { param($case, $run) $run.Workload.Memory.SampleCount }
    }
    Add-BenchmarkComparison Engine -Baseline Snapshot -Metric MedianMs
    Set-BenchmarkArtifacts Json, Csv, Markdown
}
