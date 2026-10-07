$binary = Get-BenchmarkInput Binary -Required
$repository = Get-BenchmarkInput Repository -Required
$assets = Get-BenchmarkInput Assets -Required
$evidence = Get-BenchmarkInput Evidence -Required
$rowCounts = Get-BenchmarkInput Rows 10000, 100000 -Int
$columnCounts = Get-BenchmarkInput Columns 20 -Int
$browserNames = (Get-BenchmarkInput Browsers 'Chromium,Firefox,WebKit').Split(',')
$stackNames = (Get-BenchmarkInput Stacks 'bundled,current').Split(',')
$formats = (Get-BenchmarkInput Formats 'xlsx,csv').Split(',')
$unique = Get-BenchmarkInput Unique $false -Bool

function Invoke-TableComparisonRequest {
    param([hashtable] $Request)
    $response = [DataTablesBenchmarkClient]::Current.Request(($Request | ConvertTo-Json -Compress -Depth 8)) | ConvertFrom-Json
    if (-not $response.ok) { throw $response.error }
    $response.result
}

New-BenchmarkSuite 'officeimo-browser-table-exports' -OutputRoot $evidence {
    Set-BenchmarkPolicy -Warmup 1 -Iteration 3 -Order Rotated -OutlierMode None
    Add-BenchmarkMetadata Contract 'Same grid, rows, typed numbers and literal Unicode strings; plain exports, fixed width 20, no title/footer. Browser export and control roundtrip measured; setup, transfer and independent every-value validation excluded.'
    Add-BenchmarkMetadata BinarySha256 (Get-FileHash -LiteralPath $binary -Algorithm SHA256).Hash
    Add-BenchmarkMetadata AssetManifestSha256 (Get-FileHash -LiteralPath (Join-Path $repository 'Build/BrowserExports/comparison-assets.json') -Algorithm SHA256).Hash
    Add-BenchmarkMetadata HeapBoundary 'Sampled Chromium JavaScript heap only; unavailable in Firefox/WebKit. Does not measure browser process-tree/native memory.'
    Add-BenchmarkCaseSource {
        foreach ($browser in $browserNames) { foreach ($stack in $stackNames) { foreach ($format in $formats) {
            foreach ($rows in $rowCounts) { foreach ($columns in $columnCounts) {
                [pscustomobject]@{ Name = "$browser-$stack-$format-$rows-$columns"; Browser = $browser; Stack = $stack; Format = $format; Rows = $rows; Columns = $columns; Unique = $unique }
            } }
        } } }
    }
    Set-BenchmarkSetup {
        param($case, $run)
        if ($null -eq [DataTablesBenchmarkClient]::Current) {
            [DataTablesBenchmarkClient]::Current = [DataTablesBenchmarkClient]::new($binary, $repository, (Join-Path $evidence 'browser-session'), $assets)
        }
        $run.Prepared = Invoke-TableComparisonRequest @{ command = 'prepare'; browser = $case.Browser; stack = $case.Stack; format = $case.Format; rows = $case.Rows; columns = $case.Columns; unique = $case.Unique }
        if ($run.Prepared.rows -ne $case.Rows) { throw 'Grid setup did not load the requested rows.' }
    }
    foreach ($lane in 'native','compatibility','batched') {
        Add-BenchmarkEngine $lane {
            Add-BenchmarkOperation Export {
                param($case, $run)
                $run.Measurement = Invoke-TableComparisonRequest @{ command = 'run'; lane = $case.Engine }
            }
        }
    }
    Add-BenchmarkValidation {
        param($case, $run)
        $run.Measurement | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $run.OutputDirectory 'browser-measurement.json') -Encoding utf8
        $run.Validation = Invoke-TableComparisonRequest @{ command = 'validate' }
        if (-not $run.Validation.proof.allValues -or $run.Validation.proof.rows -ne $case.Rows -or $run.Validation.proof.columns -ne $case.Columns) { throw 'Export validation differs from requested work.' }
        $run.Validation | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $run.OutputDirectory 'validation.json') -Encoding utf8
        if (-not $run.Validation.conformance.passed) { throw 'XLSX styles failed Open XML validation; measured operation and every-value proof are retained separately.' }
    }
    Add-BenchmarkMetric BrowserExportMs { param($case, $run) $run.Measurement.exportMs }
    Add-BenchmarkMetric GatherMs { param($case, $run) $run.Measurement.gatherMs }
    Add-BenchmarkMetric OutputBytes { param($case, $run) $run.Measurement.outputBytes }
    Add-BenchmarkMetric MaxTimerGapMs { param($case, $run) $run.Measurement.maxTimerGapMs }
    Add-BenchmarkMetric ValidatedCells { param($case, $run) $run.Validation.proof.cells }
    Add-BenchmarkComparison -Baseline native -Metric BrowserExportMs,OutputBytes,MaxTimerGapMs
    Set-BenchmarkArtifacts Json,Csv,Markdown
}
