$repository = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$verification = Join-Path $repository 'OfficeIMO.Project.Verification/bin/Release/net8.0/OfficeIMO.Project.Verification.dll'

New-BenchmarkSuite 'project-assess-save' -OutputRoot (Join-Path $repository 'artifacts/project/save-benchmark') {
    Set-BenchmarkPolicy -Warmup 1 -Iterations 3 -Order Rotated -OutlierMode None
    Set-BenchmarkProfile Current -Cleanup KeepOnFailure
    Add-BenchmarkMetadata 'Measurement' 'Outer duration includes process startup and input preparation. OperationMs and allocation cover assessment and edited save; construction and reopen validation are excluded.'
    Add-BenchmarkMetadata 'AffinityPolicy' 'Inherited from invoking process; keep topology, affinity, priority and power mode comparable.'
    Add-BenchmarkCaseSource @(
        [pscustomobject]@{ Name = 'Xml1k'; Count = 1000; Format = 'Xml' }
        [pscustomobject]@{ Name = 'Xml10k'; Count = 10000; Format = 'Xml' }
        [pscustomobject]@{ Name = 'Mpp14_1k'; Count = 1000; Format = 'Mpp14' }
    )
    Set-BenchmarkSetup {
        param($case, $run)
        if (!(Test-Path -LiteralPath $verification)) { throw 'Build OfficeIMO.Project.Verification in Release first.' }
    }
    Add-BenchmarkEngine OfficeIMO {
        Add-BenchmarkOperation AssessSave {
            param($case, $run)
            $json = & dotnet $verification save-scale $case.Count $case.Format
            if ($LASTEXITCODE -ne 0) { throw 'Assessment, save or independent output invariant failed.' }
            $run.Proof = $json | ConvertFrom-Json
        }
    }
    Add-BenchmarkValidation {
        param($case, $run)
        Assert-BenchmarkValue -Actual ([int]$run.Proof.tasks) -Expected ([int]$case.Count)
        if ($run.Proof.outputBytes -lt 1) { throw 'No document output was produced.' }
    }
    Add-BenchmarkMetric OperationMs { param($case, $run) $run.Proof.operationMs }
    Add-BenchmarkMetric OperationAllocatedBytes { param($case, $run) $run.Proof.allocatedBytes }
    Add-BenchmarkMetric OutputBytes { param($case, $run) $run.Proof.outputBytes }
    Add-BenchmarkMetric PeakWorkingSetBytes { param($case, $run) $run.Proof.peakWorkingSetBytes }
    Set-BenchmarkArtifacts Json, Csv, Markdown
}
