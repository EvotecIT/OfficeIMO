$repositoryRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$family = Get-BenchmarkInput Family -Required
if ($family -notin 'RenderStages', 'FirstWrite', 'ClosedLedger') { throw 'Unknown PDF phase evidence family.' }
$binaryRoot = Get-BenchmarkInput BinaryRoot -Required
$pageCounts = Get-BenchmarkInput Pages @(20, 100, 500) -Int
$rotations = Get-BenchmarkInput Rotation @(0, 90) -Int
$expectedGeometry = Get-BenchmarkInput ExpectedGeometry ''
$validationRoot = Get-BenchmarkInput ValidationRoot -Required
$textTool = Get-BenchmarkInput TextTool -Required
$textToolVersion = Get-BenchmarkInput TextToolVersion ''
$environmentEvidence = Get-BenchmarkInput EnvironmentEvidence ''
$workloads = @{}
[void] [Reflection.Assembly]::LoadFrom((Join-Path $binaryRoot 'OfficeIMO.Pdf.Benchmarks.dll'))

New-BenchmarkSuite "officeimo-pdf-$($family.ToLowerInvariant())" {
    Set-BenchmarkPolicy -Warmup 2 -Iteration 7 -Order Rotated -OutlierMode None -MemoryCleanup BeforeIteration
    Add-BenchmarkMetadata Family $family
    Add-BenchmarkMetadata Measurement 'PowerForge alone measures total elapsed duration, process-wide managed allocation, placement, rotation and raw samples. Source/intermediate preparation and artifact validation are outside each operation. No outliers are removed. These diagnostic cohorts are separate from prior ordinary runtime measurements.'
    Add-BenchmarkMetadata RenderContract 'Fixed embedded-Carlito 20-page business corpus, four 384x512 previews at scale .64. SourceOpen, Projection, Rasterization and PngEncoding prepare only their preceding public-boundary inputs outside measurement. RenderPng is the complete public call. Prepared-stage medians are not additive. Every completed preview must equal a fresh complete public control in PNG bytes and decoded pixels; this establishes managed representation consistency, not independent visual fidelity.'
    Add-BenchmarkMetadata FirstWriteContract 'Composition plus ForwardOnly Save and complete actual destination closure. Benchmark-only monotonic event timestamps observe operation-method entry, Save start, and the first successful nonempty underlying FileStream write call start/completion. Completion is an upper-bound proxy for stream/OS-cache acceptance, not first physical byte, durable persistence or delivery. The copied strategy saves to MemoryStream and ToArray before writing through the same observer. Payload length and every named noncollecting memory boundary are retained. PowerForge alone supplies total duration and a separate periodic memory sampler.'
    Add-BenchmarkMetadata GridContract 'Fresh public Load and Read(Structured) of each exact independent ReportLab closed-ledger fixture at /Rotate 0 and 90. All 4x4 cell labels remain ordered. Explicit DetectedProportions and ClosedRuling expectations qualify distinct observed geometry; cost-of-stronger-recovery comparison does not claim geometry equality.'
    if ($family -eq 'ClosedLedger') { Add-BenchmarkMetadata ExpectedGeometry $expectedGeometry }
    Add-BenchmarkMetadata MemoryContract 'Named GC.GetTotalMemory(false) and process working-set observations do not force collection and may miss peaks. Optional PowerForge sampling is a separate instrumented cohort. Neither establishes precise peak, retained heap, native allocation attribution or constant-space layout.'
    Add-BenchmarkMetadata TextToolVersion $textToolVersion
    Add-BenchmarkMetadata TextToolSha256 (Get-FileHash -LiteralPath $textTool -Algorithm SHA256).Hash
    Add-BenchmarkMetadata SpecSha256 (Get-FileHash -LiteralPath (Join-Path $PSScriptRoot 'pdf-phases.benchmark.ps1') -Algorithm SHA256).Hash
    Add-BenchmarkMetadata WrapperSha256 (Get-FileHash -LiteralPath (Join-Path $PSScriptRoot 'Run-PdfPhaseBenchmarks.ps1') -Algorithm SHA256).Hash
    foreach ($name in 'PdfRuntimeCorpus.cs', 'PdfRuntimeWorkload.cs', 'PdfRuntimeWorkload.Validation.cs', 'PdfRuntimeWorkload.Destination.cs', 'PdfRuntimeWorkload.FirstWrite.cs', 'FirstWriteFileStream.cs', 'PdfRenderStageWorkload.cs', 'PdfRenderStageWorkload.Validation.cs', 'PdfClosedLedgerRuntimeWorkload.cs') {
        Add-BenchmarkMetadata "$name.Sha256" (Get-FileHash -LiteralPath (Join-Path $repositoryRoot "OfficeIMO.Pdf.Benchmarks/$name") -Algorithm SHA256).Hash
    }
    foreach ($name in 'OfficeIMO.Core', 'OfficeIMO.Pdf', 'OfficeIMO.Pdf.Benchmarks') {
        $path = Join-Path $binaryRoot "$name.dll"
        Add-BenchmarkMetadata "$name.Sha256" (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
        $assembly = [Reflection.Assembly]::LoadFrom($path)
        Add-BenchmarkMetadata "$name.InformationalVersion" $assembly.GetCustomAttributes([Reflection.AssemblyInformationalVersionAttribute], $false)[0].InformationalVersion
    }
    $runnerRoot = Split-Path (Get-Command Invoke-BenchmarkSuite).ImplementingType.Assembly.Location
    foreach ($name in 'PSPublishModule', 'PowerForge', 'PowerForge.PowerShell') {
        Add-BenchmarkMetadata "$name.Sha256" (Get-FileHash -LiteralPath (Join-Path $runnerRoot "$name.dll") -Algorithm SHA256).Hash
    }
    if ($family -ne 'ClosedLedger') {
        Add-BenchmarkMetadata FontSource 'Existing Carlito-Regular.ttf benchmark asset, SIL Open Font License 1.1'
        Add-BenchmarkMetadata FontSha256 (Get-FileHash -LiteralPath (Join-Path $binaryRoot 'Fonts/Carlito-Regular.ttf') -Algorithm SHA256).Hash
    }
    if ($environmentEvidence) {
        Add-BenchmarkMetadata EnvironmentEvidenceSha256 (Get-FileHash -LiteralPath $environmentEvidence -Algorithm SHA256).Hash
        Add-BenchmarkMetadata EnvironmentEvidence ([IO.File]::ReadAllText($environmentEvidence))
    }
    Add-BenchmarkCaseSource {
        $selectors = if ($family -eq 'ClosedLedger') { $rotations } elseif ($family -eq 'RenderStages') { @(20) } else { $pageCounts }
        foreach ($selector in $selectors) {
            $key = "$family-$selector"
            if ($family -eq 'RenderStages') {
                $workload = [OfficeIMO.Pdf.Benchmarks.PdfRenderStageWorkload]::new()
            } elseif ($family -eq 'FirstWrite') {
                $workload = [OfficeIMO.Pdf.Benchmarks.PdfRuntimeWorkload]::new($selector)
                $workload.PrepareDestination()
            } else {
                $path = Join-Path $repositoryRoot "OfficeIMO.TestAssets/TableGeometry/ruled-ledger-$selector-native.pdf"
                $geometry = [OfficeIMO.Pdf.Benchmarks.PdfClosedLedgerExpectedGeometry] $expectedGeometry
                $workload = [OfficeIMO.Pdf.Benchmarks.PdfClosedLedgerRuntimeWorkload]::new($path, $geometry)
            }
            $workloads[$key] = $workload
            $case = [ordered]@{ Name = $key; WorkloadKey = $key; InputSha256 = $workload.SourceSha256; SourceBytes = $workload.InputBytes }
            if ($family -eq 'RenderStages') {
                $case.Pages = $workload.Pages; $case.PreviewPages = $workload.PreviewPages; $case.Scale = $workload.Scale
            } elseif ($family -eq 'FirstWrite') {
                $case.Pages = $workload.Pages; $case.ExpectedOutputSha256 = $workload.ExpectedDestinationSha256
            } else {
                $case.Pages = $workload.PageCount; $case.Rotation = $workload.Rotation
                $case.ExpectedGeometry = [string] $workload.ExpectedGeometry; $case.ManifestSha256 = $workload.ManifestSha256
                $case.ParserLimitsJson = $workload.ParserLimitsJson; $case.MaximumReadPages = $workload.MaximumReadPages
                $case.MaximumPageWorkUnits = $workload.MaximumPageWorkUnits; $case.MaximumDocumentWorkUnits = $workload.MaximumDocumentWorkUnits
            }
            [pscustomobject] $case
        }
    }
    Set-BenchmarkSetup {
        param($case, $run)
        $run.Workload = $workloads[[string] $case.WorkloadKey]
        if ($family -eq 'RenderStages') { $run.Workload.Prepare([string] $case.Operation) }
        elseif ($family -eq 'FirstWrite') { $run.Workload.PrepareFirstWrite() }
        else { $run.Workload.ReleaseResults() }
        [void] [IO.Directory]::CreateDirectory($run.OutputDirectory)
        $run.DestinationPath = Join-Path $run.OutputDirectory 'result.pdf'
        if (-not $run.Workload.HasIndependentSource) {
            $sourceDirectory = Join-Path $validationRoot "$($case.WorkloadKey)-$($case.InputSha256)"
            [void] [IO.Directory]::CreateDirectory($sourceDirectory)
            $sourcePdf = Join-Path $sourceDirectory 'source.pdf'; $sourceXml = Join-Path $sourceDirectory 'source.xhtml'
            $run.Workload.ExportSource($sourcePdf)
            & $textTool -bbox-layout -enc UTF-8 $sourcePdf $sourceXml
            if ($LASTEXITCODE -ne 0) { throw 'Independent source extraction failed.' }
            $run.Workload.ValidateIndependentSource($sourceXml)
        }
    }
    Add-BenchmarkEngine OfficeIMO {
        if ($family -eq 'RenderStages') {
            foreach ($stage in 'SourceOpen', 'Projection', 'Rasterization', 'PngEncoding', 'RenderPng') {
                Add-BenchmarkOperation $stage { param($case, $run) $run.Workload.Execute([string] $case.Operation) }
            }
        } elseif ($family -eq 'FirstWrite') {
            Add-BenchmarkOperation FirstWriteDirect { param($case, $run) $run.Workload.ExecuteObservedDestination('DestinationDirect', $run.DestinationPath, [Threading.CancellationToken]::None) }
            Add-BenchmarkOperation FirstWriteCopied { param($case, $run) $run.Workload.ExecuteObservedDestination('DestinationCopied', $run.DestinationPath, [Threading.CancellationToken]::None) }
        } else {
            Add-BenchmarkOperation ClosedLedgerRead { param($case, $run) $run.Workload.Execute() }
        }
    }
    Add-BenchmarkValidation {
        param($case, $run)
        if ($family -eq 'FirstWrite') {
            $run.Workload.ValidateFirstWrite($run.OutputDirectory, [double] $run.DurationMs)
            $xml = Join-Path $run.OutputDirectory 'result.xhtml'
            & $textTool -bbox-layout -enc UTF-8 $run.DestinationPath $xml
            if ($LASTEXITCODE -ne 0) { throw 'Independent destination extraction failed.' }
            $run.Workload.ValidateIndependentResult($xml)
        } else { $run.Workload.Validate($run.OutputDirectory) }
        $expectedPages = if ($family -eq 'RenderStages') { $case.PreviewPages } else { $case.Pages }
        Assert-BenchmarkValue -Actual $run.Workload.VerifiedPages -Expected $expectedPages
    }
    Add-BenchmarkMetric VerifiedPages { param($case, $run) $run.Workload.VerifiedPages }
    Add-BenchmarkMetric InputBytes { param($case, $run) $run.Workload.InputBytes }
    Add-BenchmarkMetric OutputBytes { param($case, $run) $run.Workload.OutputBytes }
    if ($family -eq 'RenderStages') {
        foreach ($metricName in 'SourceVerifiedPages', 'ProjectionElements', 'ProjectionTextElements', 'ProjectionImageElements', 'ProjectionShapeElements', 'CapabilityDiagnostics', 'RasterDiagnostics', 'MatchedControlPages', 'RasterBytes', 'ComparedPixelBytes') {
            Add-BenchmarkMetric $metricName { param($case, $run) $run.Workload.$metricName }
        }
    } elseif ($family -eq 'FirstWrite') {
        foreach ($metricName in 'OperationToFirstWriteCompletionMs', 'SaveToFirstWriteCompletionMs', 'OperationToFirstWriteCallStartMs', 'SaveToFirstWriteCallStartMs', 'FirstWriteCallIntervalMs', 'FirstWritePayloadBytes', 'DestinationWrittenBytes', 'DestinationNonemptyWriteCalls', 'PeakRetainedPageContentBytes', 'PeakRetainedObjectBytes', 'LargestSerializedObjectBytes') {
            Add-BenchmarkMetric $metricName { param($case, $run) $run.Workload.$metricName }
        }
        Add-BenchmarkMetric FinalByteArrayMaterialized { param($case, $run) [int] $run.Workload.FinalByteArrayMaterialized }
        foreach ($boundary in 'OperationEntry', 'CompositionComplete', 'BeforeSave', 'FirstWriteComplete', 'SaveComplete', 'DestinationClosed') {
            Add-BenchmarkMetric "$($boundary)ManagedHeapBytes" { param($case, $run) $run.Workload.PhaseMemoryBytes($boundary, 'ManagedHeap') }
            Add-BenchmarkMetric "$($boundary)WorkingSetBytes" { param($case, $run) $run.Workload.PhaseMemoryBytes($boundary, 'WorkingSet') }
        }
    } else {
        Add-BenchmarkMetric VerifiedTables { param($case, $run) $run.Workload.VerifiedTables }
        Add-BenchmarkMetric VerifiedCells { param($case, $run) $run.Workload.VerifiedCells }
    }
    Set-BenchmarkArtifacts Json, Csv, Markdown
}
