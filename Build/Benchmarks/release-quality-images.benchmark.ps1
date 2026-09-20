$repositoryRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$binaryRoot = Get-BenchmarkInput BinaryRoot (Join-Path $repositoryRoot 'OfficeIMO.Drawing.Benchmarks/bin/Release/net10.0')
if ([Environment]::Version.Major -lt 10) { throw 'Run this suite with PowerShell on .NET 10 or newer.' }

[void] [Reflection.Assembly]::LoadFrom((Join-Path $binaryRoot 'OfficeIMO.Drawing.Benchmarks.dll'))

New-BenchmarkSuite 'officeimo-release-quality-images' -OutputRoot (Join-Path $repositoryRoot '.validation/release-quality-images') {
    Set-BenchmarkPolicy -Warmup 1 -Iteration 3 -Order Rotated -OutlierMode None
    Add-BenchmarkMetadata Contract 'Provenance-bound encode, decode, metadata, optimization, and resampling evidence with fidelity, determinism, cancellation, elapsed-time, allocation, and process/native-memory metrics.'
    Add-BenchmarkMetadata BenchmarkAssemblySha256 (Get-FileHash (Join-Path $binaryRoot 'OfficeIMO.Drawing.Benchmarks.dll') -Algorithm SHA256).Hash
    Add-BenchmarkMetadata OperatingSystem ([Runtime.InteropServices.RuntimeInformation]::OSDescription)
    Add-BenchmarkMetadata Runtime ([Runtime.InteropServices.RuntimeInformation]::FrameworkDescription)
    Add-BenchmarkCaseSource {
        @(
            foreach ($item in @(
                @{ Asset = 'Screenshot'; Format = 'Png' }
                @{ Asset = 'Text'; Format = 'Jpeg' }
                @{ Asset = 'Scan'; Format = 'Tiff' }
                @{ Asset = 'AlphaGraphic'; Format = 'Webp' }
            )) {
                foreach ($operation in @('Encode', 'Decode', 'Metadata', 'Optimize')) {
                    [pscustomobject]@{
                        Name = if ($operation -eq 'Encode') { "$($item.Asset)-$($item.Format)" } else { "$($item.Asset)-$($item.Format)-$operation" }
                        Asset = $item.Asset
                        Format = $item.Format
                        Workload = $operation
                    }
                }
            }
            foreach ($asset in @('LineArt', 'Text', 'AlphaGraphic')) {
                [pscustomobject]@{ Name = "$asset-Png-Resample"; Asset = $asset; Format = 'Png'; Workload = 'Resample' }
            }
        )
    }
    Set-BenchmarkSetup {
        param($case, $run)
        $run.Workload = [OfficeIMO.Drawing.Benchmarks.ImageReleaseQualityWorkload]::new($case.Asset, $case.Format, $case.Workload)
    }
    Add-BenchmarkEngine SharedImageEngine {
        Add-BenchmarkOperation Execute { param($case, $run) $run.Workload.Execute() }
    }
    Add-BenchmarkValidation { param($case, $run) $run.Workload.Validate() }
    Add-BenchmarkMetric EncodedBytes { param($case, $run) $run.Workload.EncodedBytes }
    Add-BenchmarkMetric ManagedAllocatedBytes { param($case, $run) $run.Workload.ManagedAllocatedBytes }
    Add-BenchmarkMetric PeakWorkingSetBytes { param($case, $run) $run.Workload.PeakWorkingSetBytes }
    Add-BenchmarkMetric PeakPrivateBytes { param($case, $run) $run.Workload.PeakPrivateBytes }
    Add-BenchmarkMetric PeakNativeBytesEstimate { param($case, $run) $run.Workload.PeakNativeBytesEstimate }
    Add-BenchmarkMetric Deterministic { param($case, $run) $run.Workload.Deterministic }
    Add-BenchmarkMetric CancellationLatencyMilliseconds { param($case, $run) $run.Workload.CancellationLatencyMilliseconds }
    Add-BenchmarkMetric MeanAbsoluteError { param($case, $run) $run.Workload.MeanAbsoluteError }
    Set-BenchmarkArtifacts Json, Csv, Markdown
}
