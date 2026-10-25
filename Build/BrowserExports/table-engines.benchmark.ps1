$node = Get-BenchmarkInput Node -Required
$helper = Get-BenchmarkInput Helper -Required
$manifestPath = Get-BenchmarkInput Manifest -Required
$binary = Get-BenchmarkInput Binary -Required
$dotnet = Get-BenchmarkInput Dotnet -Required
$qpdf = Get-BenchmarkInput Qpdf -Required
$qpdfHash = Get-BenchmarkInput QpdfHash -Required
$evidence = Get-BenchmarkInput Evidence -Required
$mode = Get-BenchmarkInput Mode -Required
$owner = Get-BenchmarkInput Owner -Required
$iterations = @(Get-BenchmarkInput Iterations 3 -Int)[0]
$manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json

New-BenchmarkSuite 'officeimo-public-table-engines' -OutputRoot $evidence {
    $order = if ($mode -eq 'warm') { 'GroupedRotated' } else { 'Rotated' }
    Set-BenchmarkPolicy -Warmup 1 -Iteration $iterations -Order $order -OutlierMode None
    Add-BenchmarkMetadata Contract $manifest.contract
    Add-BenchmarkMetadata Mode $mode
    Add-BenchmarkMetadata ManifestSha256 (Get-FileHash -LiteralPath $manifestPath -Algorithm SHA256).Hash
    Add-BenchmarkMetadata IndependentReaderSha256 (Get-FileHash -LiteralPath $binary -Algorithm SHA256).Hash
    Add-BenchmarkMetadata QpdfPath $qpdf
    Add-BenchmarkMetadata QpdfSha256 $qpdfHash
    Add-BenchmarkMetadata DependencyGraphSha256 $manifest.dependencies.snapshot.sha256
    Add-BenchmarkMetadata FontBoundary 'Each PDF lane starts from the same frozen font files; file intake and per-document parsing are inside generation. Files were hashed in setup, so OS caches apply. No unrelated font intake in CSV/XLSX.'
    Add-BenchmarkMetadata DeliveryBoundary 'Complete binary buffer: includes CSV UTF-8 encoding, OfficeIMO Blob read and PDF stream concatenation. Later file copy and independent readers are excluded.'
    Add-BenchmarkMetadata MemoryBoundary 'Fresh/cold: Node lifetime peak through complete buffer includes imports/native/input. Warm: lifetime peak includes earlier exports and is not a per-export peak or cross-mode comparison. Heap categories are not additive.'
    Add-BenchmarkMetadata TimerGapBoundary 'Generation interval after module import. Cold import contributes to ApiToCompleteMs and CPU, but is outside the timer-gap sampler.'
    Add-BenchmarkMetadata WarmupBoundary 'Only PowerForge schedules warmups. Every actual warmup artifact is independently validated outside timings by the wrapper.'
    Add-BenchmarkMetadata LoadBoundary 'Normal development host; platform scheduler, no affinity/priority changes. All raw samples, load and losses retained.'
    $skippedCapabilities = @($manifest.engines.PSObject.Properties |
        Where-Object { $null -ne $_.Value.PSObject.Properties['capabilitySkip'] -and [bool] $_.Value.capabilitySkip } |
        ForEach-Object { "$($_.Name): $($_.Value.capabilitySkip)" })
    if ($skippedCapabilities.Count) { Add-BenchmarkMetadata CapabilitySkips ($skippedCapabilities -join '; ') }
    Add-BenchmarkCaseSource {
        foreach ($item in $manifest.cases) {
            [pscustomobject]@{ Name = $item.name; Format = $item.format; Rows = $item.rows; Columns = $item.columns
                Unique = $item.unique; TextProfile = $item.textProfile }
        }
    }
    Add-BenchmarkSkipRule {
        param($case)
        $engine = $manifest.engines.($case.Engine)
        return $case.Format -notin $engine.formats -or
            ($null -ne $engine.PSObject.Properties['capabilitySkip'] -and [bool] $engine.capabilitySkip)
    }
    Set-BenchmarkSetup {
        param($case, $run)
        $retained = (Get-ChildItem -LiteralPath $evidence -Filter 'sample.*' -Recurse -File | Measure-Object Length -Sum).Sum
        if ($retained -gt 512MB) { throw 'Retained comparison artifacts exceeded 512 MiB; inspect failures first.' }
        $run.ReportPath = Join-Path $run.OutputDirectory ('sample.' + $case.Format)
        # Each owner has a disjoint key prefix. Fresh/cold samples never reuse a child.
        $run.SessionKey = "$owner|$mode|$($case.Scenario)|$($case.Engine)" + $(if ($mode -ne 'warm') { "|$($run.Iteration)" })
        if (-not [DataTablesBenchmarkClient]::Sessions.ContainsKey($run.SessionKey)) {
            $arguments = [string[]] @($helper, 'session', '--manifest', $manifestPath)
            $client = [DataTablesBenchmarkClient]::new($node, $arguments, (Split-Path -Parent $helper))
            try { [DataTablesBenchmarkClient]::Sessions.Add($run.SessionKey, $client) }
            catch { $client.Dispose(); throw }
        }
        $run.ChildPid = [DataTablesBenchmarkClient]::Sessions[$run.SessionKey].ProcessId
    }
    foreach ($name in $manifest.engines.PSObject.Properties.Name) {
        Add-BenchmarkEngine $name {
            Add-BenchmarkOperation Export {
                param($case, $run)
                $request = @{ command = 'run'; engine = $case.Engine; case = $case.Scenario; mode = $mode; output = $run.ReportPath }
                try {
                    $response = [DataTablesBenchmarkClient]::Sessions[$run.SessionKey].Request(($request | ConvertTo-Json -Compress)) | ConvertFrom-Json
                    if (-not $response.ok) { throw $response.error }
                    $run.Measurement = $response.result
                    if ($run.Measurement.host.pid -ne $run.ChildPid) { throw 'Comparison child identity changed.' }
                    $run.Measurement | ConvertTo-Json -Depth 14 | Set-Content -LiteralPath (Join-Path $run.OutputDirectory 'measurement.json') -Encoding utf8
                } finally {
                    if ($mode -ne 'warm') {
                        [DataTablesBenchmarkClient]::Sessions[$run.SessionKey].Dispose()
                        [void] [DataTablesBenchmarkClient]::Sessions.Remove($run.SessionKey)
                    }
                }
            }
        }
    }
    Add-BenchmarkValidation {
        param($case, $run)
        $unique = ([bool] $case.Unique).ToString().ToLowerInvariant()
        $json = & $dotnet $binary --validate-table-engine $run.ReportPath "--format=$($case.Format)" "--rows=$($case.Rows)" `
            "--columns=$($case.Columns)" "--unique=$unique" "--text-profile=$($case.TextProfile)"
        if ($LASTEXITCODE -ne 0) { throw 'Independent table readback failed; artifact retained.' }
        $proof = $json | ConvertFrom-Json
        if (-not $proof.passed -or $proof.sha256 -ne $run.Measurement.outputSha256 -or
            $proof.rows -ne $case.Rows -or $proof.columns -ne $case.Columns -or $proof.format -ne $case.Format) {
            throw 'Independent proof does not match the measured complete artifact.'
        }
        if ($case.Format -eq 'pdf') {
            $syntaxLog = Join-Path $run.OutputDirectory 'qpdf-check.log'
            $syntaxOutput = & $qpdf --check $run.ReportPath 2>&1
            $syntaxExit = $LASTEXITCODE
            $syntaxOutput | Set-Content -LiteralPath $syntaxLog -Encoding utf8
            $syntaxProof = @{ passed = $syntaxExit -eq 0; exitCode = $syntaxExit; binarySha256 = $qpdfHash; outputSha256 = $run.Measurement.outputSha256; log = $syntaxLog }
            $syntaxProof | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $run.OutputDirectory 'qpdf-check.json') -Encoding utf8
            $proof | Add-Member -NotePropertyName qpdf -NotePropertyValue $syntaxProof
            if ($syntaxExit -ne 0) {
                $proof | ConvertTo-Json -Depth 14 | Set-Content -LiteralPath (Join-Path $run.OutputDirectory 'validation.json') -Encoding utf8
                throw 'PDF syntax check failed; artifact, reader proof and qpdf log retained.'
            }
        }
        $proof | ConvertTo-Json -Depth 14 | Set-Content -LiteralPath (Join-Path $run.OutputDirectory 'validation.json') -Encoding utf8
        Remove-Item -LiteralPath $run.ReportPath -ErrorAction Stop
        if ($mode -eq 'warm' -and $run.Iteration -eq $iterations - 1) {
            [DataTablesBenchmarkClient]::Sessions[$run.SessionKey].Dispose()
            [void] [DataTablesBenchmarkClient]::Sessions.Remove($run.SessionKey)
        }
    }
    Add-BenchmarkMetric EngineMs { param($case, $run) $run.Measurement.engineMs }
    Add-BenchmarkMetric ImportMs { param($case, $run) $run.Measurement.importMs }
    Add-BenchmarkMetric ApiToCompleteMs { param($case, $run) $run.Measurement.apiToCompleteMs }
    Add-BenchmarkMetric NodeCpuMs { param($case, $run) ($run.Measurement.cpuMicroseconds.user + $run.Measurement.cpuMicroseconds.system) / 1000 }
    Add-BenchmarkMetric ProcessLifetimePeakRssBytes { param($case, $run) $run.Measurement.peakRssBytes }
    Add-BenchmarkMetric SampledHeapBytes { param($case, $run) $run.Measurement.sampledHeapBytes }
    Add-BenchmarkMetric MaxTimerGapMs { param($case, $run) $run.Measurement.maxTimerGapMs }
    Add-BenchmarkMetric OutputBytes { param($case, $run) $run.Measurement.outputBytes }
    $metrics = @('EngineMs', 'OutputBytes')
    if ($mode -eq 'cold') { $metrics += 'ImportMs', 'ApiToCompleteMs' }
    if ($mode -ne 'warm') { $metrics += 'ProcessLifetimePeakRssBytes' }
    Add-BenchmarkComparison -Baseline officeimo -Metric $metrics
    Set-BenchmarkArtifacts Json,Csv,Markdown
}
