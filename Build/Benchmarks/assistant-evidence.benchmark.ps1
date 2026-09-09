$repositoryRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$binaryRoot = Get-BenchmarkInput BinaryRoot (Join-Path $repositoryRoot 'OfficeIMO.Studio/bin/Release/net10.0')
$pageCounts = Get-BenchmarkInput Pages 1, 100, 500 -Int
if ([Environment]::Version.Major -lt 10) { throw 'Run this suite with PowerShell on .NET 10 or newer.' }
$assemblyNames = 'OfficeIMO.Reader.Core', 'OfficeIMO.AI', 'OfficeIMO.Pdf', 'OfficeIMO.Reader.Pdf'
$references = @([IO.Directory]::GetFiles((Join-Path $PSHOME 'ref'), '*.dll'))
foreach ($name in $assemblyNames) {
    $assemblyPath = Join-Path $binaryRoot "$name.dll"
    [void] [Reflection.Assembly]::LoadFrom($assemblyPath)
    $references += $assemblyPath
}
Add-Type -Path (Join-Path $PSScriptRoot 'AssistantEvidenceWorkload.cs') -ReferencedAssemblies $references

New-BenchmarkSuite 'officeimo-assistant-evidence' -OutputRoot (Join-Path $repositoryRoot '.validation/assistant-benchmarks') {
    Set-BenchmarkPolicy -Warmup 2 -Iteration 7 -Order Rotated -OutlierMode None
    Add-BenchmarkMetadata Contract 'Same immutable PDF evidence; preparation and readiness only, excluding model inference and UI rendering.'
    Add-BenchmarkMetadata BinarySha256 (Get-FileHash (Join-Path $binaryRoot 'OfficeIMO.AI.dll') -Algorithm SHA256).Hash
    Add-BenchmarkMetadata AffinityMask ([Diagnostics.Process]::GetCurrentProcess().ProcessorAffinity.ToInt64())
    Add-BenchmarkCaseSource {
        foreach ($pages in $pageCounts) { [pscustomobject]@{ Name = "Pages-$pages"; Pages = $pages } }
    }
    Set-BenchmarkSetup {
        param($case, $run)
        $run.Workload = [AssistantEvidenceWorkload]::new($case.Pages)
    }
    Add-BenchmarkEngine Fresh {
        Add-BenchmarkOperation Readiness { param($case, $run) $run.Workload.PrepareAgain() }
    }
    Add-BenchmarkEngine Reused {
        Add-BenchmarkOperation Readiness { param($case, $run) $run.Workload.ReusePrepared() }
    }
    Add-BenchmarkValidation { param($case, $run) $run.Workload.Validate() }
    Add-BenchmarkMetric OperationAllocatedBytes { param($case, $run) $run.Workload.AllocatedBytes }
    Add-BenchmarkMetric TextCharacters { param($case, $run) $run.Workload.Readiness.TextCharacters }
    Add-BenchmarkComparison Engine -Baseline Fresh -Metric MedianMs
    Set-BenchmarkArtifacts Json, Csv, Markdown
}
