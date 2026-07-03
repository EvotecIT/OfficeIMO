param(
    [string] $OutputDirectory = "artifacts/pdf-visual-review",
    [string] $Configuration = "Debug",
    [string] $Framework = "net8.0",
    [switch] $NoRestore,
    [switch] $RequireRasterizer,
    [switch] $SkipRasterBaselines
)

$ErrorActionPreference = 'Stop'

$repoRoot = Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')
$scenarioManifestPath = Join-Path $repoRoot 'Docs/pdf-conversion-scenarios.json'
if (-not (Test-Path -LiteralPath $scenarioManifestPath)) {
    throw "PDF conversion scenario manifest was not found: $scenarioManifestPath"
}

$scenarioManifest = Get-Content -LiteralPath $scenarioManifestPath -Raw | ConvertFrom-Json
$outputPath = if ([System.IO.Path]::IsPathRooted($OutputDirectory)) {
    $OutputDirectory
} else {
    Join-Path $repoRoot $OutputDirectory
}

New-Item -ItemType Directory -Path $outputPath -Force | Out-Null
$resolvedOutputPath = (Resolve-Path -LiteralPath $outputPath).Path

$generatedReviewFileNames = @(
    'professional-report.pdf',
    'line-items-two-page.pdf',
    'headers-footers.pdf',
    'flow-dsl.pdf',
    'native-word-report.pdf',
    'native-word-daily-layout.pdf',
    'native-word-table-cell-picture-control.pdf',
    'native-powerpoint-slide.pdf',
    'native-powerpoint-dense-layout.pdf',
    'native-excel-daily-workbook.pdf',
    'excel-dashboard-report.pdf',
    'markdown-technical-document.pdf',
    'markdown-invoice-statement.pdf',
    'practical-html.pdf',
    'html-css-resource-policy.pdf',
    'html-css-resource-policy-summary.json',
    'html-pdf-roundtrip-source.pdf',
    'html-pdf-roundtrip-semantic.html',
    'html-pdf-roundtrip-positioned.html',
    'html-pdf-roundtrip-summary.json',
    'pdf-to-html-logical-source.pdf',
    'pdf-to-html-positioned-review.html',
    'pdf-logical-diagnostics-source.pdf',
    'pdf-logical-diagnostics-positioned-review.html',
    'pdf-logical-diagnostics-summary.json',
    'pdf-reader-degradation-corpus.pdf',
    'pdf-reader-degradation-summary.json',
    'pdf-reader-hostile-layout-corpus.pdf',
    'pdf-reader-hostile-layout-summary.json',
    'pdf-reader-hostile-table-corpus.pdf',
    'pdf-reader-hostile-table-summary.json',
    'powerpoint-layout-theme-groups.pdf',
    'pdf-table-import-source.pdf',
    'pdf-table-import-word.docx',
    'pdf-table-import-excel.xlsx',
    'pdf-table-import-powerpoint.pptx',
    'multilingual-business-report.pdf',
    'multilingual-word-to-pdf.pdf',
    'multilingual-excel-to-pdf.pdf',
    'multilingual-markdown-to-pdf.pdf',
    'multilingual-html-to-pdf.pdf',
    'multilingual-powerpoint-to-pdf.pdf',
    'markdown-theme-gallery-plain.pdf',
    'markdown-theme-gallery-word-like.pdf',
    'markdown-theme-gallery-technical-document.pdf',
    'markdown-theme-gallery-github-like.pdf',
    'markdown-theme-gallery-compact.pdf',
    'markdown-theme-gallery-report.pdf',
    'hello-world.pdf',
    'core-layout.pdf',
    'style-cheatsheet.pdf',
    'links-rules.pdf',
    'lists-tables.pdf',
    'table-style-gallery.pdf',
    'default-styles.pdf',
    'styled-runs.pdf',
    'tabs-leaders.pdf',
    'drawing-gallery.pdf',
    'watermark.pdf',
    'image-watermark.pdf',
    'page-border.pdf',
    'background-image.pdf',
    'background-shapes.pdf',
    'row-columns.pdf',
    'showcase-dashboard.pdf',
    'conversion-scenarios.json',
    'conversion-proof-summary.json',
    'index.md'
)

foreach ($fileName in $generatedReviewFileNames) {
    $path = Join-Path $resolvedOutputPath $fileName
    if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
    }
}

$indexPath = Join-Path $resolvedOutputPath 'index.md'

$previousReviewOutput = $env:OFFICEIMO_PDF_VISUAL_REVIEW_OUTPUT
$previousRequireRasterizer = $env:OFFICEIMO_REQUIRE_PDF_RASTERIZER
$previousSkipRasterAssertions = $env:OFFICEIMO_PDF_VISUAL_REVIEW_SKIP_RASTER_ASSERTIONS

try {
    $env:OFFICEIMO_PDF_VISUAL_REVIEW_OUTPUT = $resolvedOutputPath
    if ($RequireRasterizer) {
        $env:OFFICEIMO_REQUIRE_PDF_RASTERIZER = '1'
    } else {
        $env:OFFICEIMO_REQUIRE_PDF_RASTERIZER = $null
    }

    if ($SkipRasterBaselines) {
        $env:OFFICEIMO_PDF_VISUAL_REVIEW_SKIP_RASTER_ASSERTIONS = '1'
    } else {
        $env:OFFICEIMO_PDF_VISUAL_REVIEW_SKIP_RASTER_ASSERTIONS = $null
    }

    $testArgs = @(
        'test',
        (Join-Path $repoRoot 'OfficeIMO.Pdf.Tests/OfficeIMO.Pdf.Tests.csproj'),
        '--configuration', $Configuration,
        '--framework', $Framework,
        '--filter', 'FullyQualifiedName~PdfDocumentRasterVisualBaselineTests|FullyQualifiedName~PdfConversionScenarioManifestTests|FullyQualifiedName~PdfConversionTypographyTests',
        '--verbosity', 'minimal',
        '-p:WarningLevel=0'
    )

    if ($NoRestore) {
        $testArgs += '--no-restore'
    }

    Push-Location $repoRoot
    try {
        & dotnet @testArgs
        if ($LASTEXITCODE -ne 0) {
            throw "dotnet test failed with exit code $LASTEXITCODE."
        }
    } finally {
        Pop-Location
    }
} finally {
    $env:OFFICEIMO_PDF_VISUAL_REVIEW_OUTPUT = $previousReviewOutput
    $env:OFFICEIMO_REQUIRE_PDF_RASTERIZER = $previousRequireRasterizer
    $env:OFFICEIMO_PDF_VISUAL_REVIEW_SKIP_RASTER_ASSERTIONS = $previousSkipRasterAssertions
}

$commit = (& git -C $repoRoot rev-parse --short HEAD).Trim()
$statusLines = @(& git -C $repoRoot status --short)
$status = ($statusLines | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join [Environment]::NewLine
$generatedAt = [DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ', [Globalization.CultureInfo]::InvariantCulture)
Copy-Item -LiteralPath $scenarioManifestPath -Destination (Join-Path $resolvedOutputPath 'conversion-scenarios.json') -Force
$pdfFiles = @(
    foreach ($fileName in $generatedReviewFileNames) {
        if ($fileName -like '*.pdf') {
            $path = Join-Path $resolvedOutputPath $fileName
            if (Test-Path -LiteralPath $path) {
                Get-Item -LiteralPath $path
            }
        }
    }
) | Sort-Object Name
if ($pdfFiles.Count -eq 0) {
    throw "No PDF files were generated in $resolvedOutputPath. Check the dotnet test filter and OFFICEIMO_PDF_VISUAL_REVIEW_OUTPUT wiring."
}

$scenarioProof = @(
    foreach ($scenario in $scenarioManifest.scenarios) {
        $artifacts = @(
            foreach ($fileName in @($scenario.visualReviewFiles)) {
                $artifactPath = Join-Path $resolvedOutputPath $fileName
                if (-not (Test-Path -LiteralPath $artifactPath)) {
                    throw "Manifest scenario '$($scenario.id)' expected review artifact '$fileName', but it was not generated in $resolvedOutputPath."
                }

                $item = Get-Item -LiteralPath $artifactPath
                $hashAlgorithm = [System.Security.Cryptography.SHA256]::Create()
                try {
                    $hashBytes = $hashAlgorithm.ComputeHash([System.IO.File]::ReadAllBytes($item.FullName))
                } finally {
                    $hashAlgorithm.Dispose()
                }

                [pscustomobject]@{
                    file = $item.Name
                    sizeBytes = $item.Length
                    sha256 = (([BitConverter]::ToString($hashBytes)) -replace '-', '').ToLowerInvariant()
                }
            }
        )

        [pscustomobject]@{
            id = $scenario.id
            path = $scenario.path
            converter = $scenario.converter
            sourceFormat = $scenario.sourceFormat
            targetFormat = $scenario.targetFormat
            status = $scenario.status
            sourceFeatures = @($scenario.sourceFeatures)
            expectedSimplifications = @($scenario.expectedSimplifications)
            expectedWarnings = @($scenario.expectedWarnings)
            proof = $scenario.proof
            artifacts = $artifacts
        }
    }
)

$proofSummary = [pscustomobject]@{
    version = 1
    generatedAt = $generatedAt
    commit = $commit
    outputDirectory = $resolvedOutputPath
    manifest = 'conversion-scenarios.json'
    scenarios = $scenarioProof
}

$proofSummaryPath = Join-Path $resolvedOutputPath 'conversion-proof-summary.json'
$proofSummary | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $proofSummaryPath -Encoding UTF8

$lines = [System.Collections.Generic.List[string]]::new()
$lines.Add('# OfficeIMO PDF Visual Review Gallery')
$lines.Add('')
$lines.Add("Generated: $generatedAt")
$lines.Add('')
$lines.Add("Commit: ``$commit``")
$lines.Add('')
$lines.Add("Output directory: ``$resolvedOutputPath``")
$lines.Add('')
$lines.Add('Command:')
$lines.Add('')
$lines.Add('```powershell')
$lines.Add("Build/Export-PdfVisualReviewGallery.ps1 -OutputDirectory `"$OutputDirectory`" -Configuration `"$Configuration`" -Framework `"$Framework`"")
$lines.Add('```')
if ($SkipRasterBaselines) {
    $lines.Add('')
    $lines.Add('Raster baseline assertions were skipped for this artifact run; PDFs were still generated for review.')
}
$lines.Add('')
$lines.Add('## Review Checklist')
$lines.Add('')
$lines.Add('- Open the PDFs in Edge and Acrobat/Reader when possible.')
$lines.Add('- Check text spacing, missing glyphs, image aspect ratio, table bounds, link targets, and header/footer placement.')
$lines.Add('- Compare dense Office exports against the source intent: Word layout, Excel pagination, PowerPoint slide placement, Markdown theme rhythm, and shared PDF authoring primitives.')
$lines.Add('- Treat visual differences as product evidence, not just test noise.')
$lines.Add('')
$lines.Add('## Conversion Scenario Proof')
$lines.Add('')
$lines.Add("Manifest: [conversion-scenarios.json](conversion-scenarios.json)")
$lines.Add('')
$lines.Add("Proof summary: [conversion-proof-summary.json](conversion-proof-summary.json)")
$lines.Add('')
$lines.Add('| Scenario | Path | Converter | Review artifacts |')
$lines.Add('| --- | --- | --- | --- |')
foreach ($scenario in $scenarioProof) {
    $artifactLinks = @(
        foreach ($artifact in $scenario.artifacts) {
            $name = $artifact.file.Replace('|', '\|')
            "[$name]($name)"
        }
    ) -join ', '
    $lines.Add("| $($scenario.id) | $($scenario.path) | $($scenario.converter) | $artifactLinks |")
}
$lines.Add('')
$lines.Add('## Files')
$lines.Add('')

$lines.Add('| File | Size |')
$lines.Add('| --- | ---: |')
foreach ($file in $pdfFiles) {
    $relativeName = $file.Name.Replace('|', '\|')
    $lines.Add("| [$relativeName]($relativeName) | $($file.Length) bytes |")
}

if (-not [string]::IsNullOrWhiteSpace($status)) {
    $lines.Add('')
    $lines.Add('## Working Tree Note')
    $lines.Add('')
    $lines.Add('The repository had uncommitted changes when this gallery was generated:')
    $lines.Add('')
    $lines.Add('```text')
    $lines.Add($status)
    $lines.Add('```')
}

[System.IO.File]::WriteAllLines($indexPath, $lines, [System.Text.Encoding]::UTF8)
Write-Host "PDF visual review gallery written to $resolvedOutputPath"
Write-Host "Index: $indexPath"
