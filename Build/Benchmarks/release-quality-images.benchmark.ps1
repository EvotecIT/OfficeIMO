$repositoryRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$binaryRoot = Get-BenchmarkInput BinaryRoot (Join-Path $repositoryRoot 'OfficeIMO.Drawing.Benchmarks/bin/Release/net10.0')
if ([Environment]::Version.Major -lt 10) { throw 'Run this suite with PowerShell on .NET 10 or newer.' }

[void] [Reflection.Assembly]::LoadFrom((Join-Path $binaryRoot 'OfficeIMO.Drawing.Benchmarks.dll'))

New-BenchmarkSuite 'officeimo-release-quality-images' -OutputRoot (Join-Path $repositoryRoot '.validation/release-quality-images') {
    Set-BenchmarkPolicy -Warmup 1 -Iteration 3 -Order Rotated -OutlierMode None
    Add-BenchmarkMetadata Contract 'Deterministic bounded raster encoding with decode-based correctness, output-size, cancellation, elapsed-time, allocation, and memory evidence.'
    Add-BenchmarkMetadata BenchmarkAssemblySha256 (Get-FileHash (Join-Path $binaryRoot 'OfficeIMO.Drawing.Benchmarks.dll') -Algorithm SHA256).Hash
    Add-BenchmarkMetadata OperatingSystem ([Runtime.InteropServices.RuntimeInformation]::OSDescription)
    Add-BenchmarkMetadata Runtime ([Runtime.InteropServices.RuntimeInformation]::FrameworkDescription)
    Add-BenchmarkCaseSource {
        @(
            [pscustomobject]@{ Name = 'Screenshot-Png'; Asset = 'Screenshot'; Format = 'Png' }
            [pscustomobject]@{ Name = 'Scan-Tiff'; Asset = 'Scan'; Format = 'Tiff' }
            [pscustomobject]@{ Name = 'AlphaGraphic-Webp'; Asset = 'AlphaGraphic'; Format = 'Webp' }
            [pscustomobject]@{ Name = 'Text-Jpeg'; Asset = 'Text'; Format = 'Jpeg' }
        )
    }
    Set-BenchmarkSetup {
        param($case, $run)
        $run.Workload = [OfficeIMO.Drawing.Benchmarks.ImageReleaseQualityWorkload]::new($case.Asset, $case.Format)
    }
    Add-BenchmarkEngine SharedBoundedEncoder {
        Add-BenchmarkOperation Encode { param($case, $run) $run.Workload.Encode() }
    }
    Add-BenchmarkValidation { param($case, $run) $run.Workload.Validate() }
    Add-BenchmarkMetric EncodedBytes { param($case, $run) $run.Workload.EncodedBytes }
    Add-BenchmarkMetric ManagedAllocatedBytes { param($case, $run) $run.Workload.ManagedAllocatedBytes }
    Add-BenchmarkMetric Deterministic { param($case, $run) $run.Workload.Deterministic }
    Add-BenchmarkMetric CancellationObserved { param($case, $run) $run.Workload.CancellationObserved }
    Add-BenchmarkMetric MeanAbsoluteError { param($case, $run) $run.Workload.MeanAbsoluteError }
    Set-BenchmarkArtifacts Json, Csv, Markdown
}
