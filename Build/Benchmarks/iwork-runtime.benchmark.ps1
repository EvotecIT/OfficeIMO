$repositoryRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$binaryRoot = Get-BenchmarkInput BinaryRoot (Join-Path $repositoryRoot 'OfficeIMO.IWork.Benchmarks/bin/Release/net8.0')
$factors = Get-BenchmarkInput Factors 1, 10, 100 -Int
$measureRetainedMemory = Get-BenchmarkInput MeasureRetainedMemory $false -Bool
$nativeCopyCase = Get-BenchmarkInput NativeCopyCase $false -Bool
$specPath = Join-Path $PSScriptRoot 'iwork-runtime.benchmark.ps1'
$workloads = @{}
$retainedBaselines = @{}
[void] [Reflection.Assembly]::LoadFrom((Join-Path $binaryRoot 'OfficeIMO.IWork.Benchmarks.dll'))

New-BenchmarkSuite 'officeimo-iwork-runtime' {
    Set-BenchmarkPolicy -Warmup 2 -Iteration 5 -Order Rotated -OutlierMode None -MemoryCleanup BeforeIteration
    Add-BenchmarkMetadata Contract 'Deterministic synthetic ZIP or pinned native Pages/Numbers/Keynote fixture: load/project or load/convert/save. Cancellation cases request cancellation after actual package-read or native-copy I/O. Input/model preparation and full semantic readback are outside timing. No native Apple appearance qualification.'
    Add-BenchmarkMetadata Measurement 'PowerForge elapsed time and managed allocation including host invocation. Working-set deltas are not peak or retained memory. Saving uses normal owner APIs.'
    Add-BenchmarkMetadata RetainedMemory 'Opt-in collected process-wide managed heap before execution and after validation/result release. Includes host/cache effects; excludes native and peak memory.'
    Add-BenchmarkMetadata SampledMemory 'Opt-in PowerForge managed-heap/resident observations during operations. Maxima are observed lower bounds, include observer/host effects and do not isolate native allocations. Instrumented timing uses a separate run mode.'
    Add-BenchmarkMetadata AffinityPolicy 'Inherited; macOS processor placement is unqualified. Keep competing work idle and record host power mode.'
    Add-BenchmarkMetadata Runtime ([Runtime.InteropServices.RuntimeInformation]::FrameworkDescription)
    Add-BenchmarkMetadata RunnerSha256 (Get-FileHash ([PowerForge.PowerShellBenchmarkRunner].Assembly.Location) -Algorithm SHA256).Hash
    Add-BenchmarkMetadata WorkloadSha256 (Get-FileHash (Join-Path $binaryRoot 'OfficeIMO.IWork.Benchmarks.dll') -Algorithm SHA256).Hash
    Add-BenchmarkMetadata SpecSha256 (Get-FileHash $specPath -Algorithm SHA256).Hash
    foreach ($name in 'OfficeIMO.Core', 'OfficeIMO.IWork', 'OfficeIMO.Word', 'OfficeIMO.Word.IWork', 'OfficeIMO.Excel', 'OfficeIMO.Excel.IWork', 'OfficeIMO.PowerPoint', 'OfficeIMO.PowerPoint.IWork') {
        Add-BenchmarkMetadata "$name.Sha256" (Get-FileHash (Join-Path $binaryRoot "$name.dll") -Algorithm SHA256).Hash
    }
    Add-BenchmarkCaseSource {
        if ($nativeCopyCase) {
            $workloads['Keynote-1'] = [OfficeIMO.IWork.Benchmarks.IWorkRuntimeWorkload]::new('Keynote', 1)
            $workloads['Keynote-1'].Prepare('CancelDuringNativeCopy')
            [pscustomobject]@{ Name = 'Keynote-Native'; Kind = 'Keynote'; Source = 'OwnedNativeCopy'; ConversionPolicy = 'CancelledNativeStreamCopy'; MeasureRetainedMemory = $measureRetainedMemory; Units = 1; InputSha256 = $workloads['Keynote-1'].NativeCopyInputSha256 }
            return
        }
        foreach ($factor in $factors) {
            $size = switch ($factor) { 1 { 'Small' } 10 { 'Medium' } 100 { 'Large' } 0 { 'Native' } default { throw 'Unknown iWork workload scale.' } }
            foreach ($kind in 'Pages', 'Numbers', 'Keynote') {
                $base = switch ($kind) { Pages { 100 } Numbers { 1000 } Keynote { 10 } }
                $units = if ($factor -eq 0) { switch ($kind) { Pages { 3 } Numbers { 9 } Keynote { 2 } } } else { $base * $factor }
                $key = "$kind-$units"
                $workloads[$key] = if ($factor -eq 0) {
                    [OfficeIMO.IWork.Benchmarks.IWorkRuntimeWorkload]::new($kind, (Join-Path $repositoryRoot 'OfficeIMO.TestAssets/Documents/IWorkCorpus'))
                } else { [OfficeIMO.IWork.Benchmarks.IWorkRuntimeWorkload]::new($kind, $units) }
                [pscustomobject]@{ Name = "$kind-$size"; Kind = $kind; Source = $size; ConversionPolicy = 'CompleteEditable'; MeasureRetainedMemory = $measureRetainedMemory; Units = $units; InputSha256 = $workloads[$key].InputSha256 }
            }
        }
    }
    Set-BenchmarkSetup {
        param($case, $run)
        $key = "$($case.Kind)-$($case.Units)"
        if (-not $workloads.ContainsKey($key)) {
            $workloads[$key] = if ($case.Source -eq 'Native') {
                [OfficeIMO.IWork.Benchmarks.IWorkRuntimeWorkload]::new($case.Kind, (Join-Path $repositoryRoot 'OfficeIMO.TestAssets/Documents/IWorkCorpus'))
            } else { [OfficeIMO.IWork.Benchmarks.IWorkRuntimeWorkload]::new($case.Kind, $case.Units) }
        }
        $run.Workload = $workloads[$key]
        $run.Workload.Prepare($case.Operation)
        if ($measureRetainedMemory) {
            $run.MemoryProbe = [PowerForge.BenchmarkManagedMemoryProbe]::new()
        }
    }
    Add-BenchmarkEngine OfficeIMO {
        foreach ($operation in 'LoadProject', 'ConvertSave', 'CancelDuringLoad', 'CancelDuringConvert', 'CancelDuringNativeCopy') {
            Add-BenchmarkOperation $operation {
                param($case, $run)
                $run.Workload.Execute($case.Operation)
            }
        }
    }
    Add-BenchmarkValidation {
        param($case, $run)
        $run.Workload.Validate()
        $expected = if ($case.Operation -like 'CancelDuring*') { 1 } else { $case.Units }
        Assert-BenchmarkValue -Actual $run.Workload.VerifiedUnits -Expected $expected
        $run.OutputBytes = $run.Workload.OutputBytes
        $run.Workload.ReleaseResults()
        if ($measureRetainedMemory) {
            $run.RetainedMemory = $run.MemoryProbe.Capture()
            $key = "$($case.Kind)-$($case.Source)-$($case.Units)-$($case.Operation)"
            if (-not $retainedBaselines.ContainsKey($key)) { $retainedBaselines[$key] = $run.RetainedMemory.BaselineBytes }
            $run.RepeatedManagedBaseline = $retainedBaselines[$key]
        }
    }
    Add-BenchmarkMetric VerifiedUnits { param($case, $run) $run.Workload.VerifiedUnits }
    Add-BenchmarkMetric InputBytes { param($case, $run) $run.Workload.InputBytes }
    Add-BenchmarkMetric OutputBytes { param($case, $run) $run.OutputBytes }
    Add-BenchmarkMetric CancellationObserved { param($case, $run) [int]$run.Workload.CancellationObserved }
    Add-BenchmarkMetric CancellationProcessedBytes { param($case, $run) $run.Workload.CancellationProcessedBytes }
    Add-BenchmarkMetric CancellationTotalIoBytes { param($case, $run) $run.Workload.CancellationTotalIoBytes }
    Add-BenchmarkMetric CancellationLatencyMs { param($case, $run) $run.Workload.CancellationLatencyMs }
    if ($measureRetainedMemory) {
        Add-BenchmarkMetric ManagedBaselineBytes { param($case, $run) $run.RetainedMemory.BaselineBytes }
        Add-BenchmarkMetric CollectedManagedBytes { param($case, $run) $run.RetainedMemory.CollectedBytes }
        Add-BenchmarkMetric RetainedManagedDeltaBytes { param($case, $run) $run.RetainedMemory.DeltaBytes }
        Add-BenchmarkMetric RepeatedManagedBaselineBytes { param($case, $run) $run.RepeatedManagedBaseline }
        Add-BenchmarkMetric RepeatedRetainedManagedDeltaBytes { param($case, $run) $run.RetainedMemory.CollectedBytes - $run.RepeatedManagedBaseline }
    }
    Set-BenchmarkArtifacts Json, Csv, Markdown
}
