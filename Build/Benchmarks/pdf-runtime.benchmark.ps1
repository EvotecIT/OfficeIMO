$repositoryRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$binaryRoot = Get-BenchmarkInput BinaryRoot (Join-Path $repositoryRoot 'OfficeIMO.Pdf.Benchmarks/bin/Release/net10.0')
$pageCounts = Get-BenchmarkInput Pages 1, 20, 100 -Int
$validationRoot = Get-BenchmarkInput ValidationRoot (Join-Path $repositoryRoot 'Ignore/Benchmarks/PdfRuntime/inputs')
$textTool = Get-BenchmarkInput TextTool 'pdftotext'
$textToolVersion = Get-BenchmarkInput TextToolVersion ''
$environmentEvidence = Get-BenchmarkInput EnvironmentEvidence ''
$specPath = Join-Path $PSScriptRoot 'pdf-runtime.benchmark.ps1'
$wrapperPath = Join-Path $PSScriptRoot 'Run-PdfRuntimeBenchmarks.ps1'
$workloads = @{}
[void] [Reflection.Assembly]::LoadFrom((Join-Path $binaryRoot 'OfficeIMO.Pdf.Benchmarks.dll'))

New-BenchmarkSuite 'officeimo-pdf-runtime' {
    Set-BenchmarkPolicy -Warmup 2 -Iteration 7 -Order Rotated -OutlierMode None -MemoryCleanup BeforeIteration
    Add-BenchmarkMetadata Contract 'pdf-runtime-v1: explicit 600 x 800 point pages. BusinessEmbeddedCarlito inputs contain embedded regular Carlito text, numeric table values, a colored rectangle and one repeated 32 x 32 image. TextEditStandard14 inputs contain four ordinary Helvetica paragraphs per page, without table, rectangle or image paint. Region edits use only TextEditStandard14; all other operations use BusinessEmbeddedCarlito. Generation includes composition, font registration, full layout, serialization and MemoryStream.ToArray. Font file reading is prepared outside timing. Region edits touch one prepared middle-page region, include source loading and saved serialization, and preserve neighboring words.'
    Add-BenchmarkMetadata AnalysisContract 'InspectCold loads four documents; InspectReused loads one document and inspects it four times within the operation. Neither lane reuses a document across measured iterations.'
    Add-BenchmarkMetadata RenderContract 'SVG: first min(pages, 12) pages, including default OpenType glyph aliases and nearest-neighbor image vector rectangles. PNG: first min(pages, 4) pages at 512 maximum dimension, 384 x 512 pixels. Complete public Render.Pages call is measured. Artifact decoding/readback is outside timing.'
    Add-BenchmarkMetadata WriterContract 'Buffered uses PDF 1.4; Forward uses PDF 1.7. Both compress content, save to MemoryStream and copy to byte[]. Forward means object serialization, not complete pipeline streaming.'
    Add-BenchmarkMetadata ValidationContract 'Managed page/text/writer and decoded-render checks on every measured sample; Poppler bbox extraction validates every PDF page and retained neighboring word geometry. SVG retains default OpenType shaping and validates logical drawing text plus byte-exact public ToDrawing/export consistency, including embedded painted-glyph cmap aliases. PNG validates decoded dimensions and visible content. Rendering checks establish managed representation consistency, not independent visual fidelity.'
    Add-BenchmarkMetadata Measurement 'PowerForge owns elapsed timing, process-wide managed allocation, rotation and raw samples. Setup, source preparation, readback, export and validation are excluded. Instrumented memory runs are separate cohorts.'
    Add-BenchmarkMetadata TextToolPath $textTool
    Add-BenchmarkMetadata TextToolVersion $textToolVersion
    Add-BenchmarkMetadata TextToolSha256 (Get-FileHash -LiteralPath $textTool -Algorithm SHA256).Hash
    Add-BenchmarkMetadata FontSource 'Website/Apps/OfficeIMO.Web.Converter/Assets/Fonts/Carlito-Regular.ttf; existing benchmark-project asset, SIL Open Font License 1.1'
    Add-BenchmarkMetadata FontSha256 (Get-FileHash -LiteralPath (Join-Path $binaryRoot 'Fonts/Carlito-Regular.ttf') -Algorithm SHA256).Hash
    Add-BenchmarkMetadata SpecSha256 (Get-FileHash -LiteralPath $specPath -Algorithm SHA256).Hash
    Add-BenchmarkMetadata WrapperSha256 (Get-FileHash -LiteralPath $wrapperPath -Algorithm SHA256).Hash
    foreach ($name in 'PdfRuntimeCorpus.cs', 'PdfRuntimeWorkload.cs', 'PdfRuntimeWorkload.Validation.cs') {
        Add-BenchmarkMetadata "$name.Sha256" (Get-FileHash -LiteralPath (Join-Path $repositoryRoot "OfficeIMO.Pdf.Benchmarks/$name") -Algorithm SHA256).Hash
    }
    if ($environmentEvidence) {
        Add-BenchmarkMetadata EnvironmentEvidenceSha256 (Get-FileHash -LiteralPath $environmentEvidence -Algorithm SHA256).Hash
        Add-BenchmarkMetadata EnvironmentEvidence ([IO.File]::ReadAllText($environmentEvidence))
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
    Add-BenchmarkCaseSource {
        foreach ($pages in $pageCounts) {
            $key = "Pages-$pages"
            $business = [OfficeIMO.Pdf.Benchmarks.PdfRuntimeWorkload]::new($pages, 'BusinessEmbeddedCarlito')
            $editing = [OfficeIMO.Pdf.Benchmarks.PdfRuntimeWorkload]::new($pages, 'TextEditStandard14')
            $workloads["Business-$key"] = $business
            $workloads["Editing-$key"] = $editing
            [pscustomobject]@{
                Name = $key; WorkloadKey = $key; Pages = $pages; TargetPage = $business.TargetPage
                BusinessSourceSha256 = $business.SourceSha256; EditSourceSha256 = $editing.SourceSha256
                SourceBytes = $business.InputBytes; EditSourceBytes = $editing.InputBytes
                SvgPages = [Math]::Min($pages, 12); PngPages = [Math]::Min($pages, 4)
                InspectionCalls = 4; MoveDeltaX = 8; MoveDeltaY = 0
            }
        }
    }
    Set-BenchmarkSetup {
        param($case, $run)
        $profile = if ($case.Operation -in 'ReplaceRegion', 'MoveRegion') { 'Editing' } else { 'Business' }
        $run.Workload = $workloads["$profile-$($case.WorkloadKey)"]
        $run.Workload.ReleaseResults()
        if (-not $run.Workload.HasIndependentSource) {
            $sourceDirectory = Join-Path $validationRoot "$profile-$($case.WorkloadKey)-$($run.Workload.SourceSha256)"
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
        Add-BenchmarkOperation GenerateBuffered { param($case, $run) $run.Workload.Execute('GenerateBuffered') }
        Add-BenchmarkOperation GenerateForward { param($case, $run) $run.Workload.Execute('GenerateForward') }
        Add-BenchmarkOperation InspectCold { param($case, $run) $run.Workload.Execute('InspectCold') }
        Add-BenchmarkOperation InspectReused { param($case, $run) $run.Workload.Execute('InspectReused') }
        Add-BenchmarkOperation RenderSvg { param($case, $run) $run.Workload.Execute('RenderSvg') }
        Add-BenchmarkOperation RenderPng { param($case, $run) $run.Workload.Execute('RenderPng') }
        Add-BenchmarkOperation ReplaceRegion { param($case, $run) $run.Workload.Execute('ReplaceRegion') }
        Add-BenchmarkOperation MoveRegion { param($case, $run) $run.Workload.Execute('MoveRegion') }
    }
    Add-BenchmarkValidation {
        param($case, $run)
        $run.Workload.Validate()
        $run.Workload.ExportEvidence($run.OutputDirectory)
        if ($run.Workload.HasPdfOutput) {
            $pdf = Join-Path $run.OutputDirectory 'result.pdf'
            $xml = Join-Path $run.OutputDirectory 'result.xhtml'
            & $textTool -bbox-layout -enc UTF-8 $pdf $xml
            if ($LASTEXITCODE -ne 0) { throw 'Independent saved-output extraction failed.' }
            $run.Workload.ValidateIndependentResult($xml)
        }
        $expectedPages = switch ($case.Operation) {
            'RenderSvg' { $case.SvgPages }
            'RenderPng' { $case.PngPages }
            default { $case.Pages }
        }
        Assert-BenchmarkValue -Actual $run.Workload.VerifiedPages -Expected $expectedPages
    }
    Add-BenchmarkMetric VerifiedPages { param($case, $run) $run.Workload.VerifiedPages }
    Add-BenchmarkMetric InputBytes { param($case, $run) $run.Workload.InputBytes }
    Add-BenchmarkMetric OutputBytes { param($case, $run) $run.Workload.OutputBytes }
    Add-BenchmarkMetric AffectedTextSpans { param($case, $run) $run.Workload.AffectedTextSpans }
    Add-BenchmarkMetric PeakRetainedPageContentBytes { param($case, $run) $run.Workload.PeakRetainedPageContentBytes }
    Add-BenchmarkMetric PeakRetainedObjectBytes { param($case, $run) $run.Workload.PeakRetainedObjectBytes }
    Add-BenchmarkMetric LargestSerializedObjectBytes { param($case, $run) $run.Workload.LargestSerializedObjectBytes }
    Set-BenchmarkArtifacts Json, Csv, Markdown
}
