$node = Get-BenchmarkInput Node -Required
$helper = Get-BenchmarkInput Helper -Required
$manifest = Get-BenchmarkInput Manifest -Required
$qpdf = Get-BenchmarkInput Qpdf -Required
$pdftotext = Get-BenchmarkInput Pdftotext -Required
$python = Get-BenchmarkInput Python -Required
$evidence = Get-BenchmarkInput Evidence -Required
$compare = Get-BenchmarkInput Compare $true -Bool
$prepared = Get-Content -LiteralPath $manifest -Raw | ConvertFrom-Json
$cases = $prepared.cases
$engines = @($prepared.engines.PSObject.Properties.Name)
$comparisonBaselines = @((Get-BenchmarkInput ComparisonBaselines ($prepared.comparisonBaselines -join ',')) -split ',' | Where-Object { $_ })
if (@($comparisonBaselines | Where-Object { $_ -eq 'current' -or $engines -cnotcontains $_ }).Count -gt 0) { throw 'Comparison baselines must name declared reference engines.' }
$equivalentBaselineCases = @($cases | Where-Object { $_.PSObject.Properties['equivalentBaselineContract'] } | ForEach-Object name)

New-BenchmarkSuite 'officeimo-complete-report-workloads' -OutputRoot $evidence {
    Set-BenchmarkPolicy -Warmup 1 -Iteration 3 -Order Rotated -OutlierMode None
    Add-BenchmarkMetadata Contract 'Deterministic inventory, assessment and grouped service-analysis reports. Lazy row construction is included; import/setup/validation excluded. Every cell is independently read back; tagged PDF also checks the pypdf-resolved logical tree and marked-content ownership. Editable/grouped analysis cases and wide cases without an explicit matching-baseline declaration have no baseline ranking. Declared matching wide baseline/current cohorts require the same complete bands, repeated identifiers and logical values. Grouped PDF is untagged.'
    if ($equivalentBaselineCases.Count -gt 0) {
        Add-BenchmarkMetadata EquivalentBaselineCases ($equivalentBaselineCases -join ',')
    }
    Add-BenchmarkMetadata DeliveryBoundary 'Sink cases include awaited regular-file writes (OS page cache applies). Blob cases time complete Blob generation and exclude subsequent Blob read/file copying. Do not compare sink and Blob lanes as equivalent memory contracts.'
    Add-BenchmarkMetadata FirstDeliveryMs 'First accepted sink write for sink cases. For Blob cases, completed Blob availability; raw firstByteMs is null because no earlier bytes are exposed.'
    Add-BenchmarkMetadata MemoryBoundary 'Fresh Node lifetime peakRSS through operation end includes imports/native memory. Sampled JS heap and external memory can miss synchronous peaks and are not additive. PowerForge host allocation metrics describe PowerShell, not the Node engine.'
    Add-BenchmarkMetadata ExecutionBoundary 'Each sample is a fresh Node process with its module imported before EngineMs and fresh JIT state. PowerForge warmup processes exit before measured processes. This is module-preloaded/JIT-fresh export latency, not persistent steady-state warmed export latency.'
    Add-BenchmarkMetadata LoadBoundary 'Normal development host and platform scheduler; no affinity or priority changes. Per-sample load averages are retained. Rotated warmups/iterations and all losses remain visible.'
    Add-BenchmarkMetadata ManifestSha256 (Get-FileHash -LiteralPath $manifest -Algorithm SHA256).Hash
    Add-BenchmarkMetadata ComparisonEnabled $compare
    Add-BenchmarkMetadata RuntimeSources ($prepared.engines | ConvertTo-Json -Compress -Depth 4)
    Add-BenchmarkCaseSource { foreach ($fixture in $cases) { [pscustomobject]@{ Name = $fixture.name; Format = $fixture.format } } }
    Set-BenchmarkSetup {
        param($case, $run)
        $retained = (Get-ChildItem -LiteralPath $evidence -Filter 'sample.*' -Recurse -File | Measure-Object Length -Sum).Sum
        if ($retained -gt 512MB) { throw 'Retained report artifacts exceeded 512 MiB; inspect validation before generating more.' }
        $run.ReportPath = Join-Path $run.OutputDirectory ('sample.' + $case.Format)
    }
    foreach ($lane in $engines) {
        Add-BenchmarkEngine $lane {
            Add-BenchmarkOperation Export {
                param($case, $run)
                $json = & $node $helper run --manifest $manifest --case $case.Scenario --engine $case.Engine --output $run.ReportPath
                if ($LASTEXITCODE -ne 0) { throw 'Report generation failed; inspect the Node error.' }
                $run.Measurement = $json | ConvertFrom-Json
                $run.Measurement | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $run.OutputDirectory 'measurement.json') -Encoding utf8
            }
        }
    }
    Add-BenchmarkValidation {
        param($case, $run)
        $json = & $node $helper validate --manifest $manifest --case $case.Scenario --output $run.ReportPath --qpdf $qpdf --pdftotext $pdftotext --python $python
        if ($LASTEXITCODE -ne 0) { throw 'Independent report readback failed; sample retained.' }
        $run.Validation = $json | ConvertFrom-Json
        if (-not $run.Validation.passed -or $run.Validation.sha256 -ne $run.Measurement.outputSha256) { throw 'Validated report differs from the measured output.' }
        $run.Validation | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $run.OutputDirectory 'validation.json') -Encoding utf8
        Remove-Item -LiteralPath $run.ReportPath -ErrorAction Stop
    }
    Add-BenchmarkMetric EngineMs { param($case, $run) $run.Measurement.engineMs }
    Add-BenchmarkMetric RowsPerSecond { param($case, $run) 1000 * $run.Measurement.rows / $run.Measurement.engineMs }
    Add-BenchmarkMetric NodeCpuMs { param($case, $run) ($run.Measurement.cpuMicroseconds.user + $run.Measurement.cpuMicroseconds.system) / 1000 }
    Add-BenchmarkMetric FirstDeliveryMs { param($case, $run) if ($null -eq $run.Measurement.firstByteMs) { $run.Measurement.engineMs } else { $run.Measurement.firstByteMs } }
    Add-BenchmarkMetric NodePeakRssBytes { param($case, $run) $run.Measurement.peakRssBytes }
    Add-BenchmarkMetric SampledHeapBytes { param($case, $run) $run.Measurement.maxSampledHeapBytes }
    Add-BenchmarkMetric MaxTimerGapMs { param($case, $run) $run.Measurement.maxTimerGapMs }
    Add-BenchmarkMetric OutputBytes { param($case, $run) $run.Measurement.outputBytes }
    if ($compare) {
        foreach ($baseline in $comparisonBaselines) {
            Add-BenchmarkComparison -Baseline $baseline -Metric EngineMs,FirstDeliveryMs,NodePeakRssBytes,SampledHeapBytes,MaxTimerGapMs,OutputBytes
        }
    }
    Set-BenchmarkArtifacts Json,Csv,Markdown
}
