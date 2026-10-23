$binary = Get-BenchmarkInput Binary -Required
$repository = Get-BenchmarkInput Repository -Required
$assets = Get-BenchmarkInput Assets -Required
$evidence = Get-BenchmarkInput Evidence -Required
$rowCounts = Get-BenchmarkInput Rows 10000, 100000 -Int
$columnCounts = Get-BenchmarkInput Columns 20 -Int
$browserNames = (Get-BenchmarkInput Browsers 'Chromium,Firefox,WebKit').Split(',')
$stackNames = (Get-BenchmarkInput Stacks 'bundled,current').Split(',')
$formats = (Get-BenchmarkInput Formats 'csv').Split(',')
$unique = Get-BenchmarkInput Unique $false -Bool
$styled = Get-BenchmarkInput Styled $false -Bool
$fullWidthScan = Get-BenchmarkInput FullWidthScan $false -Bool
$textProfile = Get-BenchmarkInput TextProfile 'unicode'
$comparison = Get-BenchmarkInput Comparison $true -Bool
if ($formats.Where({ $_ -notin 'csv','xlsx','pdf' }).Count) { throw 'Supported formats are csv, xlsx and pdf.' }
if ($styled -and $formats.Where({ $_ -ne 'pdf' }).Count) { throw 'The styled comparison currently supports PDF only.' }
if ($comparison -and $formats -contains 'xlsx' -and -not $fullWidthScan) { throw 'XLSX comparisons require full-width scanning in both implementations.' }

function Invoke-TableComparisonRequest {
    param([hashtable] $Request)
    $response = [DataTablesBenchmarkClient]::Current.Request(($Request | ConvertTo-Json -Compress -Depth 8)) | ConvertFrom-Json
    if (-not $response.ok) { throw $response.error }
    $response.result
}

# Shared output ownership: read the completed Blob before the next prepare can replace it.
# Successful PowerForge warmups bypass Add-BenchmarkValidation, so they call this same block
# after the browser has recorded exportMs. Measured requests use the normal validation callback.
$validateComparisonExport = {
    param($case, $run)
    try {
        $run.Measurement | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $run.OutputDirectory 'browser-measurement.json') -Encoding utf8
        # This shared block also runs from an operation's guarded scope. Keep its request self-contained.
        $response = [DataTablesBenchmarkClient]::Current.Request('{"command":"validate","keep":true}') | ConvertFrom-Json
        if (-not $response.ok) { throw $response.error }
        $run.Validation = $response.result
        if (-not $run.Validation.proof.allValues -or $run.Validation.proof.rows -ne $case.Rows -or $run.Validation.proof.columns -ne $case.Columns) { throw 'Export validation differs from requested work.' }
        $run.Validation | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $run.OutputDirectory 'validation.json') -Encoding utf8
        if (-not $run.Validation.conformance.passed) { throw 'XLSX styles failed Open XML validation; measured operation and every-value proof are retained separately.' }
        if ($case.Format -eq 'pdf') {
            if ([string]::IsNullOrWhiteSpace($run.Validation.artifactPath) -or -not (Test-Path -LiteralPath $run.Validation.artifactPath -PathType Leaf)) { throw 'The actual PDF was not retained for independent qpdf validation.' }
            $qpdf = (Get-Command qpdf -CommandType Application -ErrorAction Stop).Source
            & $qpdf --check $run.Validation.artifactPath *> (Join-Path $run.OutputDirectory 'qpdf-check.txt')
            if ($LASTEXITCODE -ne 0) { throw 'The actual browser PDF failed qpdf validation.' }
            $run.Validation | Add-Member -NotePropertyName qpdf -NotePropertyValue $true
            $run.Validation | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $run.OutputDirectory 'validation.json') -Encoding utf8
        }
    } catch {
        # This wrapper owns Current. Do not leave a failed warmup's browser/session alive.
        try {
            if ($null -ne [DataTablesBenchmarkClient]::Current) {
                [DataTablesBenchmarkClient]::Current.Dispose()
            }
        } finally { [DataTablesBenchmarkClient]::Current = $null }
        throw
    }
}

New-BenchmarkSuite 'officeimo-browser-table-exports' -OutputRoot $evidence {
    Set-BenchmarkPolicy -Warmup 1 -Iteration 3 -Order Rotated -OutlierMode None
    Add-BenchmarkMetadata Contract 'Same grid and ordered values; quoted CSV, typed XLSX or font-supported PDF text, no title/footer. Browser export and control roundtrip measured; setup, transfer and independent cell validation excluded. Comparative XLSX scans all rows for approximate widths; fixed-width XLSX is output qualification only.'
    Add-BenchmarkMetadata ComparisonEnabled $comparison
    Add-BenchmarkMetadata TextProfile $textProfile
    Add-BenchmarkMetadata PdfContract 'Identical Carlito regular/bold fonts, A3 landscape, 6-point text, equal columns, repeated heading, same line spacing and 0.4-point grid. Optional body bands. Generation timed through completed Blob, not early Buttons callback; raw font bytes prepared before timing, per-document font parsing included in generation for both writers. Each rendered value and column position validated independently.'
    Add-BenchmarkMetadata XlsxContract 'FullWidthScan measures complete-table character sizing in both implementations, bounded at width 54. Approximation algorithms and resulting widths can differ. OfficeIMO explicitly budgets retained serialized rows for this workload; this is not its default bounded leading sample or a streaming-memory comparison.'
    Add-BenchmarkMetadata BinarySha256 (Get-FileHash -LiteralPath $binary -Algorithm SHA256).Hash
    Add-BenchmarkMetadata AssetManifestSha256 (Get-FileHash -LiteralPath (Join-Path $repository 'Build/BrowserExports/comparison-assets.json') -Algorithm SHA256).Hash
    Add-BenchmarkMetadata HeapBoundary 'Sampled Chromium JavaScript heap only; unavailable in Firefox/WebKit. Does not measure browser process-tree/native memory.'
    Add-BenchmarkCaseSource {
        foreach ($browser in $browserNames) { foreach ($stack in $stackNames) { foreach ($format in $formats) {
            foreach ($rows in $rowCounts) { foreach ($columns in $columnCounts) {
                [pscustomobject]@{ Name = "$browser-$stack-$format-$rows-$columns"; Browser = $browser; Stack = $stack; Format = $format; Rows = $rows; Columns = $columns; Unique = $unique; Styled = $styled; FullWidthScan = $fullWidthScan; TextProfile = $textProfile }
            } }
        } } }
    }
    Set-BenchmarkSetup {
        param($case, $run)
        if ($null -eq [DataTablesBenchmarkClient]::Current) {
            [DataTablesBenchmarkClient]::Current = [DataTablesBenchmarkClient]::new($binary, $repository, (Join-Path $evidence 'browser-session'), $assets)
        }
        $run.Prepared = Invoke-TableComparisonRequest @{ command = 'prepare'; browser = $case.Browser; stack = $case.Stack; format = $case.Format; rows = $case.Rows; columns = $case.Columns; unique = $case.Unique; styled = $case.Styled; fullWidthScan = $case.FullWidthScan; textProfile = $case.TextProfile }
        if ($run.Prepared.rows -ne $case.Rows) { throw 'Grid setup did not load the requested rows.' }
    }
    foreach ($lane in 'native','compatibility','batched') {
        Add-BenchmarkEngine $lane {
            Add-BenchmarkOperation Export {
                param($case, $run)
                $run.Measurement = Invoke-TableComparisonRequest @{ command = 'run'; lane = $case.Engine }
                if ($run.Iteration -lt 0) { & $validateComparisonExport $case $run }
            }
        }
    }
    Add-BenchmarkValidation $validateComparisonExport
    Add-BenchmarkMetric BrowserExportMs { param($case, $run) $run.Measurement.exportMs }
    Add-BenchmarkMetric GatherMs { param($case, $run) $run.Measurement.gatherMs }
    Add-BenchmarkMetric OutputBytes { param($case, $run) $run.Measurement.outputBytes }
    Add-BenchmarkMetric MaxTimerGapMs { param($case, $run) $run.Measurement.maxTimerGapMs }
    Add-BenchmarkMetric ValidatedCells { param($case, $run) $run.Validation.proof.cells }
    if ($comparison) { Add-BenchmarkComparison -Baseline native -Metric BrowserExportMs,OutputBytes,MaxTimerGapMs }
    Set-BenchmarkArtifacts Json,Csv,Markdown
}
