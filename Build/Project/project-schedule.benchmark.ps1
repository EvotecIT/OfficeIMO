$repository = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$verification = Join-Path $repository 'OfficeIMO.Project.Verification/bin/Release/net8.0/OfficeIMO.Project.Verification.dll'

New-BenchmarkSuite 'project-schedule' -OutputRoot (Join-Path $repository 'artifacts/project-bundle2/schedule-benchmark') {
    Set-BenchmarkPolicy -Warmup 1 -Iterations 3 -Order Rotated -OutlierMode None
    Set-BenchmarkProfile Current -Cleanup KeepOnFailure
    Add-BenchmarkMetadata 'Measurement' 'Process startup, graph creation, date calculation, and independent chain-result validation.'
    Add-BenchmarkMetadata 'Processor' ([Environment]::GetEnvironmentVariable('PROCESSOR_IDENTIFIER'))
    Add-BenchmarkMetadata 'AffinityPolicy' 'Inherited from invoking process; use the same domain for all cases and retain host topology and power context.'
    Add-BenchmarkCaseSource @(
        [pscustomobject]@{ Name = 'Dense1k'; Count = 1000 }
        [pscustomobject]@{ Name = 'Dense10k'; Count = 10000 }
        [pscustomobject]@{ Name = 'Dense100k'; Count = 100000 }
    )
    Set-BenchmarkSetup {
        param($case, $run)
        if (!(Test-Path -LiteralPath $verification)) { throw 'Build OfficeIMO.Project.Verification in Release first.' }
    }
    Add-BenchmarkEngine OfficeIMO {
        Add-BenchmarkOperation CreateCalculateValidate {
            param($case, $run)
            $json = & dotnet $verification schedule-scale $case.Count dense
            if ($LASTEXITCODE -ne 0) { throw 'Schedule calculation or independent chain validation failed.' }
            $run.Proof = $json | ConvertFrom-Json
        }
    }
    Add-BenchmarkValidation {
        param($case, $run)
        Assert-BenchmarkValue -Actual ([int]$run.Proof.results) -Expected ([int]$case.Count)
        if ($run.Proof.peakWorkingSetBytes -gt 2GB -or $run.Proof.allocatedBytes -gt 4GB) { throw 'Schedule memory budget exceeded.' }
    }
    Add-BenchmarkMetric CalculationAllocatedBytes { param($case, $run) $run.Proof.allocatedBytes }
    Add-BenchmarkMetric PeakWorkingSetBytes { param($case, $run) $run.Proof.peakWorkingSetBytes }
    Set-BenchmarkArtifacts Json, Csv, Markdown
}
