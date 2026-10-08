$repositoryRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$binaryRoot = Get-BenchmarkInput BinaryRoot (Join-Path $repositoryRoot 'OfficeIMO.Pdf.Benchmarks/bin/Release/net8.0')
$pageCounts = Get-BenchmarkInput Pages 1, 25, 100 -Int
$specPath = Join-Path $PSScriptRoot 'pdf-redaction-runtime.benchmark.ps1'
$workloads = @{}
[void] [Reflection.Assembly]::LoadFrom((Join-Path $binaryRoot 'OfficeIMO.Pdf.Benchmarks.dll'))

New-BenchmarkSuite 'officeimo-pdf-redaction-runtime' {
    Set-BenchmarkPolicy -Warmup 2 -Iteration 5 -Order Rotated -OutlierMode None -MemoryCleanup BeforeIteration
    Add-BenchmarkMetadata Contract 'Precise regex search, relabeled review, apply and complete-stream/managed-render verification on immutable synthetic PDFs. Input generation and final saved-output readback are outside measurement. No UI rendering or cross-engine ranking.'
    Add-BenchmarkMetadata Runtime ([Runtime.InteropServices.RuntimeInformation]::FrameworkDescription)
    Add-BenchmarkMetadata AffinityPolicy 'Inherited. Keep other builds idle; macOS processor placement is unqualified.'
    Add-BenchmarkMetadata WorkloadSha256 (Get-FileHash (Join-Path $binaryRoot 'OfficeIMO.Pdf.Benchmarks.dll') -Algorithm SHA256).Hash
    Add-BenchmarkMetadata PdfSha256 (Get-FileHash (Join-Path $binaryRoot 'OfficeIMO.Pdf.dll') -Algorithm SHA256).Hash
    Add-BenchmarkMetadata RunnerModuleSha256 (Get-FileHash (Get-Command Invoke-BenchmarkSuite).ImplementingType.Assembly.Location -Algorithm SHA256).Hash
    Add-BenchmarkMetadata SpecSha256 (Get-FileHash $specPath -Algorithm SHA256).Hash
    Add-BenchmarkCaseSource {
        foreach ($pages in $pageCounts) {
            $workloads[$pages] = [OfficeIMO.Pdf.Benchmarks.PdfRedactionRuntimeWorkload]::new($pages)
            [pscustomobject]@{ Name = "Pages-$pages"; Pages = $pages; InputSha256 = $workloads[$pages].SourceSha256 }
        }
    }
    Set-BenchmarkSetup {
        param($case, $run)
        $run.Workload = $workloads[[int]$case.Pages]
    }
    Add-BenchmarkEngine OfficeIMO {
        Add-BenchmarkOperation SearchReviewApply { param($case, $run) $run.Workload.Execute() }
    }
    Add-BenchmarkValidation {
        param($case, $run)
        $run.Workload.Validate()
        Assert-BenchmarkValue -Actual $run.Workload.VerifiedPages -Expected $case.Pages
        Assert-BenchmarkValue -Actual $run.Workload.SelectedAreas -Expected $case.Pages
        $run.VerifiedPages = $run.Workload.VerifiedPages
        $run.InputBytes = $run.Workload.InputBytes
        $run.OutputBytes = $run.Workload.OutputBytes
        $run.Workload.ReleaseResults()
    }
    Add-BenchmarkMetric VerifiedPages { param($case, $run) $run.VerifiedPages }
    Add-BenchmarkMetric InputBytes { param($case, $run) $run.InputBytes }
    Add-BenchmarkMetric OutputBytes { param($case, $run) $run.OutputBytes }
    Set-BenchmarkArtifacts Json, Csv, Markdown
}
