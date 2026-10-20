$node = Get-BenchmarkInput Node -Required
$helper = Get-BenchmarkInput Helper -Required
$manifest = Get-BenchmarkInput Manifest -Required
$qpdf = Get-BenchmarkInput Qpdf -Required
$pdftotext = Get-BenchmarkInput Pdftotext -Required
$evidence = Get-BenchmarkInput Evidence -Required
$dimensions = Get-BenchmarkInput Dimensions 2048,4096 -Int
$includeFonts = Get-BenchmarkInput IncludeFonts $true -Bool

New-BenchmarkSuite 'officeimo-engine-resources' -OutputRoot $evidence {
    Set-BenchmarkPolicy -Warmup 1 -Iteration 3 -Order Rotated -OutlierMode None
    Add-BenchmarkMetadata Contract 'Immutable baseline/current OfficeIMO JavaScript. Every output is byte-identical to baseline and independently checked. PNG cases measure complete uncompressed PDF Blob generation; font cases measure constructor versus PdfFont.create only, with later PDF generation/validation excluded.'
    Add-BenchmarkMetadata ProcessBoundary 'Fresh Node child per operation. Module/input loading excluded from EngineMs; fresh-process JIT and existing OS file caches apply. PowerForge wall/allocation metrics describe its PowerShell host and include child startup; compare EngineMs for the engine operation.'
    Add-BenchmarkMetadata MemoryBoundary 'Node process-lifetime peakRSS through operation end includes startup/input/native memory; not process tree. HeapUsed/external/arrayBuffers deltas are snapshots and overlapping categories, not peak totals. Font PDF generation happens after captured memory metrics.'
    Add-BenchmarkMetadata LoadBoundary 'Normal development Mac under recorded per-sample load. No affinity pin; platform scheduler. Defaults warmup1/iterations3; actual counts recorded by shared runner. Rotated ordering, no outlier removal; losses and all raw samples retained.'
    Add-BenchmarkMetadata Manifest $manifest
    Add-BenchmarkMetadata HelperSha256 (Get-FileHash -LiteralPath $helper -Algorithm SHA256).Hash
    Add-BenchmarkCaseSource {
        foreach ($dimension in $dimensions) { [pscustomobject]@{ Name = "png-$dimension"; Kind = 'png'; Dimension = $dimension } }
        if ($includeFonts) {
            [pscustomobject]@{ Name = 'font-carlito'; Kind = 'font'; Dimension = 0 }
            [pscustomobject]@{ Name = 'font-japanese'; Kind = 'font'; Dimension = 0 }
        }
    }
    Set-BenchmarkSetup {
        param($case, $run)
        $retained = (Get-ChildItem -LiteralPath $evidence -Filter sample.pdf -Recurse -File | Measure-Object Length -Sum).Sum
        if ($retained -gt 300MB) { throw 'Retained failed PDFs reached300MiB; inspect failures before generating more.' }
        $run.PdfPath = Join-Path $run.OutputDirectory 'sample.pdf'
    }
    foreach ($lane in 'baseline','current') {
        Add-BenchmarkEngine $lane {
            Add-BenchmarkOperation Export {
                param($case, $run)
                $json = & $node $helper run --manifest $manifest --case $case.Scenario --engine $case.Engine --output $run.PdfPath
                if ($LASTEXITCODE -ne 0) { throw 'Resource operation failed; inspect the Node error.' }
                $run.Measurement = $json | ConvertFrom-Json
                $run.Measurement | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $run.OutputDirectory 'measurement.json') -Encoding utf8
            }
        }
    }
    Add-BenchmarkValidation {
        param($case, $run)
        $json = & $node $helper validate --manifest $manifest --case $case.Scenario --output $run.PdfPath --qpdf $qpdf --pdftotext $pdftotext
        if ($LASTEXITCODE -ne 0) { throw 'Independent resource output validation failed; sample PDF retained.' }
        $run.Validation = $json | ConvertFrom-Json
        $run.Validation | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $run.OutputDirectory 'validation.json') -Encoding utf8
        if (-not $run.Validation.passed -or $run.Validation.sha256 -ne $run.Measurement.outputSha256) { throw 'Validation does not match the measured complete PDF.' }
        Remove-Item -LiteralPath $run.PdfPath -ErrorAction Stop
    }
    Add-BenchmarkMetric EngineMs { param($case, $run) $run.Measurement.engineMs }
    Add-BenchmarkMetric MaxTimerGapMs { param($case, $run) $run.Measurement.maxTimerGapMs }
    Add-BenchmarkMetric NodePeakRssBytes { param($case, $run) $run.Measurement.peakRssBytes }
    Add-BenchmarkMetric HeapUsedDelta { param($case, $run) $run.Measurement.heapUsedDelta }
    Add-BenchmarkMetric ExternalDelta { param($case, $run) $run.Measurement.externalDelta }
    Add-BenchmarkMetric ArrayBuffersDelta { param($case, $run) $run.Measurement.arrayBuffersDelta }
    Add-BenchmarkMetric OutputBytes { param($case, $run) $run.Measurement.outputBytes }
    Add-BenchmarkComparison -Baseline baseline -Metric EngineMs,MaxTimerGapMs,NodePeakRssBytes,OutputBytes
    Set-BenchmarkArtifacts Json,Csv,Markdown
}
