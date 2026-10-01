$repositoryRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$binaryRoot = Get-BenchmarkInput BinaryRoot (Join-Path $repositoryRoot 'OfficeIMO.IWork.Benchmarks/bin/Release/net8.0')
$factors = Get-BenchmarkInput Factors 1, 10, 100 -Int
$specPath = Join-Path $PSScriptRoot 'iwork-runtime.benchmark.ps1'
$workloads = @{}
[void] [Reflection.Assembly]::LoadFrom((Join-Path $binaryRoot 'OfficeIMO.IWork.Benchmarks.dll'))

New-BenchmarkSuite 'officeimo-iwork-runtime' {
    Set-BenchmarkPolicy -Warmup 2 -Iteration 5 -Order Rotated -OutlierMode None -MemoryCleanup BeforeIteration
    Add-BenchmarkMetadata Contract 'Deterministic synthetic supported-format ZIP: load/project or load/convert/save. Input generation and full semantic readback are outside timing. No native Apple appearance qualification.'
    Add-BenchmarkMetadata Measurement 'PowerForge elapsed time and managed allocation including host invocation. Working-set deltas are not peak or retained memory. Saving uses normal owner APIs.'
    Add-BenchmarkMetadata AffinityPolicy 'Inherited; macOS processor placement is unqualified. Keep competing work idle and record host power mode.'
    Add-BenchmarkMetadata Runtime ([Runtime.InteropServices.RuntimeInformation]::FrameworkDescription)
    Add-BenchmarkMetadata RunnerSha256 (Get-FileHash ([PowerForge.PowerShellBenchmarkRunner].Assembly.Location) -Algorithm SHA256).Hash
    Add-BenchmarkMetadata WorkloadSha256 (Get-FileHash (Join-Path $binaryRoot 'OfficeIMO.IWork.Benchmarks.dll') -Algorithm SHA256).Hash
    Add-BenchmarkMetadata SpecSha256 (Get-FileHash $specPath -Algorithm SHA256).Hash
    foreach ($name in 'OfficeIMO.Core', 'OfficeIMO.IWork', 'OfficeIMO.Word', 'OfficeIMO.Word.IWork', 'OfficeIMO.Excel', 'OfficeIMO.Excel.IWork', 'OfficeIMO.PowerPoint', 'OfficeIMO.PowerPoint.IWork') {
        Add-BenchmarkMetadata "$name.Sha256" (Get-FileHash (Join-Path $binaryRoot "$name.dll") -Algorithm SHA256).Hash
    }
    Add-BenchmarkCaseSource {
        foreach ($factor in $factors) {
            $size = switch ($factor) { 1 { 'Small' } 10 { 'Medium' } 100 { 'Large' } default { throw 'Unknown iWork workload scale.' } }
            foreach ($kind in 'Pages', 'Numbers', 'Keynote') {
                $base = switch ($kind) { Pages { 100 } Numbers { 1000 } Keynote { 10 } }
                $units = $base * $factor
                $key = "$kind-$units"
                $workloads[$key] = [OfficeIMO.IWork.Benchmarks.IWorkRuntimeWorkload]::new($kind, $units)
                [pscustomobject]@{ Name = "$kind-$size"; Kind = $kind; Units = $units; InputSha256 = $workloads[$key].InputSha256 }
            }
        }
    }
    Set-BenchmarkSetup {
        param($case, $run)
        $key = "$($case.Kind)-$($case.Units)"
        if (-not $workloads.ContainsKey($key)) {
            $workloads[$key] = [OfficeIMO.IWork.Benchmarks.IWorkRuntimeWorkload]::new($case.Kind, $case.Units)
        }
        $run.Workload = $workloads[$key]
    }
    Add-BenchmarkEngine OfficeIMO {
        foreach ($operation in 'LoadProject', 'ConvertSave') {
            Add-BenchmarkOperation $operation {
                param($case, $run)
                $run.Workload.Execute($case.Operation)
            }
        }
    }
    Add-BenchmarkValidation {
        param($case, $run)
        $run.Workload.Validate()
        Assert-BenchmarkValue -Actual $run.Workload.VerifiedUnits -Expected $case.Units
    }
    Add-BenchmarkMetric VerifiedUnits { param($case, $run) $run.Workload.VerifiedUnits }
    Add-BenchmarkMetric InputBytes { param($case, $run) $run.Workload.InputBytes }
    Add-BenchmarkMetric OutputBytes { param($case, $run) $run.Workload.OutputBytes }
    Set-BenchmarkArtifacts Json, Csv, Markdown
}
