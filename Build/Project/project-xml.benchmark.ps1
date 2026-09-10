$repository = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$verification = Join-Path $repository 'OfficeIMO.Project.Verification/bin/Release/net8.0/OfficeIMO.Project.Verification.dll'
$inputs = Join-Path $repository 'artifacts/project/scale-inputs'

New-BenchmarkSuite 'project-xml-lifecycle' -OutputRoot (Join-Path $repository 'artifacts/project/scale-results') {
    Set-BenchmarkPolicy -Warmup 1 -Iterations 3 -Order Rotated -OutlierMode None
    Set-BenchmarkProfile Current -Cleanup KeepOnFailure
    Add-BenchmarkMetadata 'Measurement' 'Process startup, bounded load, validation, scalar edit, save, and independent document readback; input generation is excluded.'
    Add-BenchmarkMetadata 'Processor' ([Environment]::GetEnvironmentVariable('PROCESSOR_IDENTIFIER'))
    Add-BenchmarkMetadata 'MedianRuntimeBudgetMs' 30000
    Add-BenchmarkMetadata 'AffinityPolicy' 'Inherited from invoking process; record actual affinity, topology, priority, and power mode with evidence.'
    Add-BenchmarkCaseSource @(
        [pscustomobject]@{ Name = 'Flat1k'; Count = 1000; Shape = 'flat' }
        [pscustomobject]@{ Name = 'Flat10k'; Count = 10000; Shape = 'flat' }
        [pscustomobject]@{ Name = 'Flat100k'; Count = 100000; Shape = 'flat' }
        [pscustomobject]@{ Name = 'Deep10k'; Count = 10000; Shape = 'deep' }
        [pscustomobject]@{ Name = 'Dense10k'; Count = 10000; Shape = 'dense' }
        [pscustomobject]@{ Name = 'Timephased10k'; Count = 10000; Shape = 'timephased' }
    )
    Set-BenchmarkSetup {
        param($case, $run)
        if (!(Test-Path -LiteralPath $verification)) { throw 'Build OfficeIMO.Project.Verification in Release before running.' }
        [void][IO.Directory]::CreateDirectory($inputs)
        $run.Input = Join-Path $inputs ($case.Scenario + '.xml')
        if (!(Test-Path -LiteralPath $run.Input)) {
            & dotnet $verification scale-create $run.Input $case.Count $case.Shape | Out-Null
            if ($LASTEXITCODE -ne 0) { throw 'Workload input generation failed.' }
        }
    }
    Add-BenchmarkEngine OfficeIMO {
        Add-BenchmarkOperation ReadEditSave {
            param($case, $run)
            $json = & dotnet $verification scale-read-edit-save $run.Input $case.Count $case.Shape
            if ($LASTEXITCODE -ne 0) { throw 'Project lifecycle workload failed.' }
            $run.Proof = $json | ConvertFrom-Json
        }
    }
    Add-BenchmarkValidation {
        param($case, $run)
        Assert-BenchmarkValue -Actual ([long]$run.Proof.tasks) -Expected ([long]$case.Count)
        Assert-BenchmarkValue -Actual ([long]$run.Proof.assignments) -Expected ([long]$case.Count)
        Assert-BenchmarkValue -Actual ([long]$run.Proof.uidSum) -Expected ([long]([long]$case.Count * ([long]$case.Count + 1) / 2))
        if ($run.Proof.peakWorkingSetBytes -gt 2GB) { throw 'Lifecycle workload exceeded the 2 GiB process budget.' }
        if ($run.Proof.allocatedBytes -gt 4GB) { throw 'Load/edit/save exceeded the 4 GiB allocation budget.' }
    }
    Add-BenchmarkMetric InputBytes { param($case, $run) $run.Proof.inputBytes }
    Add-BenchmarkMetric OutputBytes { param($case, $run) $run.Proof.outputBytes }
    Add-BenchmarkMetric LoadSaveAllocatedBytes { param($case, $run) $run.Proof.allocatedBytes }
    Add-BenchmarkMetric PeakWorkingSetBytes { param($case, $run) $run.Proof.peakWorkingSetBytes }
    Set-BenchmarkArtifacts Json, Csv, Markdown
}
