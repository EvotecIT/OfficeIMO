$repositoryRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$family = Get-BenchmarkInput Family 'StructuredRead'
if ($family -notin 'StructuredRead', 'Destination') { throw 'Unknown PDF evidence family.' }
$binaryRoot = Get-BenchmarkInput BinaryRoot (Join-Path $repositoryRoot 'OfficeIMO.Pdf.Benchmarks/bin/Release/net10.0')
$pageCounts = Get-BenchmarkInput Pages $(if ($family -eq 'StructuredRead') { 20, 100, 500 } else { 1, 20, 100 }) -Int
$validationRoot = Get-BenchmarkInput ValidationRoot (Join-Path $repositoryRoot 'Ignore/Benchmarks/PdfRuntimeEvidence/inputs')
$textTool = Get-BenchmarkInput TextTool 'pdftotext'
$textToolVersion = Get-BenchmarkInput TextToolVersion ''
$environmentEvidence = Get-BenchmarkInput EnvironmentEvidence ''
$specPath = Join-Path $PSScriptRoot 'pdf-runtime-evidence.benchmark.ps1'
$workloads = @{}
[void] [Reflection.Assembly]::LoadFrom((Join-Path $binaryRoot 'OfficeIMO.Pdf.Benchmarks.dll'))

New-BenchmarkSuite "officeimo-pdf-$($family.ToLowerInvariant())" {
    Set-BenchmarkPolicy -Warmup 2 -Iteration 7 -Order Rotated -OutlierMode None -MemoryCleanup BeforeIteration
    Add-BenchmarkMetadata Family $family
    Add-BenchmarkMetadata Measurement 'PowerForge owns elapsed timing, process-wide managed allocation, placement, rotation and raw samples; outliers are retained. Source/reference preparation, readback, artifact qualification and export are outside measurement. Sampled memory is a separate cohort.'
    Add-BenchmarkMetadata StructuredContract 'Fresh Load(byte[], default parser limits).Read(Structured, default built-in work budgets, CancellationToken.None). Every 600 x 800 point page has Running operations at x=36+.02*p, top=36, font=10+.001*p; Body markerD4 at x=36, top=240, font=12. p is one-based. Validate exactly one header per page and the complete body sequence.'
    Add-BenchmarkMetadata DestinationContract 'Both lanes compose the identical embedded Carlito business corpus with PDF 1.7, compressed streams and ForwardOnly object serialization. DestinationDirect saves to FileStream; DestinationCopied saves to MemoryStream, calls ToArray and File.WriteAllBytes. Both close a complete file inside measurement. No forced durable flush; OS cache/storage effects are part of the result. Complete forward layout is retained by both. CancellationToken.None is passed to Save; no cancellation latency is measured.'
    Add-BenchmarkMetadata ValidationContract 'Poppler independently qualifies every prepared PDF page. Structured read validates logical header classification and exact ordered body content. Every destination is read after measurement, checked against an exact prepared forward-PDF digest, qualified by managed page/text/writer assertions and independently extracted through Poppler. Digests are integrity proof, not independent producer proof.'
    Add-BenchmarkMetadata TextToolPath $textTool
    Add-BenchmarkMetadata TextToolVersion $textToolVersion
    Add-BenchmarkMetadata TextToolSha256 (Get-FileHash -LiteralPath $textTool -Algorithm SHA256).Hash
    Add-BenchmarkMetadata SpecSha256 (Get-FileHash -LiteralPath $specPath -Algorithm SHA256).Hash
    Add-BenchmarkMetadata WrapperSha256 (Get-FileHash -LiteralPath (Join-Path $PSScriptRoot 'Run-PdfRuntimeEvidenceBenchmarks.ps1') -Algorithm SHA256).Hash
    foreach ($name in 'PdfRuntimeCorpus.cs', 'PdfRuntimeWorkload.cs', 'PdfRuntimeWorkload.Validation.cs', 'PdfRuntimeWorkload.Destination.cs', 'PdfStructuredReadRuntimeWorkload.cs') {
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
    if ($family -eq 'Destination') {
        Add-BenchmarkMetadata FontSource 'Existing Website/Apps/OfficeIMO.Web.Converter/Assets/Fonts/Carlito-Regular.ttf; SIL Open Font License 1.1'
        Add-BenchmarkMetadata FontSha256 (Get-FileHash -LiteralPath (Join-Path $binaryRoot 'Fonts/Carlito-Regular.ttf') -Algorithm SHA256).Hash
    }
    if ($environmentEvidence) {
        Add-BenchmarkMetadata EnvironmentEvidenceSha256 (Get-FileHash -LiteralPath $environmentEvidence -Algorithm SHA256).Hash
        Add-BenchmarkMetadata EnvironmentEvidence ([IO.File]::ReadAllText($environmentEvidence))
    }
    Add-BenchmarkCaseSource {
        foreach ($pages in $pageCounts) {
            $key = "$family-Pages-$pages"
            if ($family -eq 'StructuredRead') {
                $workload = [OfficeIMO.Pdf.Benchmarks.PdfStructuredReadRuntimeWorkload]::new($pages)
            } else {
                $workload = [OfficeIMO.Pdf.Benchmarks.PdfRuntimeWorkload]::new($pages)
                $workload.PrepareDestination()
            }
            $workloads[$key] = $workload
            $case = [ordered]@{ Name = $key; WorkloadKey = $key; Pages = $pages; InputSha256 = $workload.SourceSha256; SourceBytes = $workload.InputBytes }
            if ($family -eq 'StructuredRead') {
                $case.MaximumReadPages = $workload.MaximumReadPages
                $case.MaximumPageWorkUnits = $workload.MaximumPageWorkUnits
                $case.MaximumDocumentWorkUnits = $workload.MaximumDocumentWorkUnits
                $case.ParserLimitsJson = $workload.ParserLimitsJson
            } else { $case.ExpectedOutputSha256 = $workload.ExpectedDestinationSha256 }
            [pscustomobject] $case
        }
    }
    Set-BenchmarkSetup {
        param($case, $run)
        $run.Workload = $workloads[[string]$case.WorkloadKey]
        $run.Workload.ReleaseResults()
        [void] [IO.Directory]::CreateDirectory($run.OutputDirectory)
        $run.DestinationPath = Join-Path $run.OutputDirectory 'result.pdf'
        if (-not $run.Workload.HasIndependentSource) {
            $sourceDirectory = Join-Path $validationRoot "$($case.WorkloadKey)-$($case.InputSha256)"
            [void] [IO.Directory]::CreateDirectory($sourceDirectory)
            $sourcePdf = Join-Path $sourceDirectory 'source.pdf'
            $sourceXml = Join-Path $sourceDirectory 'source.xhtml'
            $run.Workload.ExportSource($sourcePdf)
            & $textTool -bbox-layout -enc UTF-8 $sourcePdf $sourceXml
            if ($LASTEXITCODE -ne 0) { throw 'Independent source extraction failed.' }
            $run.Workload.ValidateIndependentSource($sourceXml)
        }
    }
    Add-BenchmarkEngine OfficeIMO {
        if ($family -eq 'StructuredRead') {
            Add-BenchmarkOperation StructuredRead { param($case, $run) $run.Workload.Execute([Threading.CancellationToken]::None) }
        } else {
            Add-BenchmarkOperation DestinationDirect { param($case, $run) $run.Workload.ExecuteDestination('DestinationDirect', $run.DestinationPath, [Threading.CancellationToken]::None) }
            Add-BenchmarkOperation DestinationCopied { param($case, $run) $run.Workload.ExecuteDestination('DestinationCopied', $run.DestinationPath, [Threading.CancellationToken]::None) }
        }
    }
    Add-BenchmarkValidation {
        param($case, $run)
        if ($family -eq 'StructuredRead') {
            $run.Workload.Validate($run.OutputDirectory)
        } else {
            $run.Workload.ValidateDestination($run.OutputDirectory)
            $pdf = Join-Path $run.OutputDirectory 'result.pdf'
            $xml = Join-Path $run.OutputDirectory 'result.xhtml'
            & $textTool -bbox-layout -enc UTF-8 $pdf $xml
            if ($LASTEXITCODE -ne 0) { throw 'Independent destination extraction failed.' }
            $run.Workload.ValidateIndependentResult($xml)
        }
        Assert-BenchmarkValue -Actual $run.Workload.VerifiedPages -Expected $case.Pages
    }
    Add-BenchmarkMetric VerifiedPages { param($case, $run) $run.Workload.VerifiedPages }
    Add-BenchmarkMetric InputBytes { param($case, $run) $run.Workload.InputBytes }
    if ($family -eq 'StructuredRead') {
        Add-BenchmarkMetric VerifiedHeaders { param($case, $run) $run.Workload.VerifiedHeaders }
        Add-BenchmarkMetric VerifiedBodyMarkers { param($case, $run) $run.Workload.VerifiedBodyMarkers }
    } else {
        Add-BenchmarkMetric OutputBytes { param($case, $run) $run.Workload.OutputBytes }
        Add-BenchmarkMetric FinalByteArrayMaterialized { param($case, $run) [int] $run.Workload.FinalByteArrayMaterialized }
        Add-BenchmarkMetric PeakRetainedPageContentBytes { param($case, $run) $run.Workload.PeakRetainedPageContentBytes }
        Add-BenchmarkMetric PeakRetainedObjectBytes { param($case, $run) $run.Workload.PeakRetainedObjectBytes }
        Add-BenchmarkMetric LargestSerializedObjectBytes { param($case, $run) $run.Workload.LargestSerializedObjectBytes }
        Add-BenchmarkMetric PageContentMemoryLimitBytes { param($case, $run) $run.Workload.PageContentMemoryLimitBytes }
        Add-BenchmarkMetric ObjectBufferMemoryLimitBytes { param($case, $run) $run.Workload.ObjectBufferMemoryLimitBytes }
        Add-BenchmarkMetric PageContentSpilled { param($case, $run) [int] $run.Workload.PageContentSpilled }
        Add-BenchmarkMetric ObjectBufferSpilled { param($case, $run) [int] $run.Workload.ObjectBufferSpilled }
    }
    Set-BenchmarkArtifacts Json, Csv, Markdown
}
