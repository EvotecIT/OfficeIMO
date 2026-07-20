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

$manifestReviewFileNames = @(
    foreach ($scenario in @($scenarioManifest.scenarios)) {
        foreach ($fileName in @($scenario.visualReviewFiles) + @($scenario.sourceReviewFiles)) {
            if (-not [string]::IsNullOrWhiteSpace($fileName)) {
                $fileName
            }
        }
    }
) | Sort-Object -Unique

$standaloneReviewFileNames = @(
    'professional-report.pdf',
    'line-items-two-page.pdf',
    'headers-footers.pdf',
    'flow-dsl.pdf',
    'native-word-daily-layout.pdf',
    'native-word-table-cell-picture-control.pdf',
    'native-powerpoint-slide.pdf',
    'multilingual-business-report.pdf',
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
    'pdf-conversion-support-matrix.md',
    'index.md'
)

$generatedReviewFileNames = @(
    $manifestReviewFileNames
    $standaloneReviewFileNames
) | Sort-Object -Unique

foreach ($fileName in $generatedReviewFileNames) {
    $path = Join-Path $resolvedOutputPath $fileName
    if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
    }
}

$indexPath = Join-Path $resolvedOutputPath 'index.md'
$supportMatrixPath = Join-Path $resolvedOutputPath 'pdf-conversion-support-matrix.md'

$previousReviewOutput = $env:OFFICEIMO_PDF_VISUAL_REVIEW_OUTPUT
$previousRequireRasterizer = $env:OFFICEIMO_REQUIRE_PDF_RASTERIZER
$previousRequireRasterBaselineMatch = $env:OFFICEIMO_REQUIRE_PDF_RASTER_BASELINE_MATCH
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
        $env:OFFICEIMO_REQUIRE_PDF_RASTER_BASELINE_MATCH = $null
    } else {
        $env:OFFICEIMO_PDF_VISUAL_REVIEW_SKIP_RASTER_ASSERTIONS = $null
        $env:OFFICEIMO_REQUIRE_PDF_RASTER_BASELINE_MATCH = '1'
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

        $oneNoteTestArgs = @(
            'test',
            (Join-Path $repoRoot 'OfficeIMO.OneNote.Tests/OfficeIMO.OneNote.Tests.csproj'),
            '--configuration', $Configuration,
            '--framework', $Framework,
            '--filter', 'FullyQualifiedName~OneNotePdfVisualScenarioTests',
            '--verbosity', 'minimal',
            '-p:WarningLevel=0'
        )
        if ($NoRestore) {
            $oneNoteTestArgs += '--no-restore'
        }
        & dotnet @oneNoteTestArgs
        if ($LASTEXITCODE -ne 0) {
            throw "OneNote PDF visual scenario test failed with exit code $LASTEXITCODE."
        }
    } finally {
        Pop-Location
    }
} finally {
    $env:OFFICEIMO_PDF_VISUAL_REVIEW_OUTPUT = $previousReviewOutput
    $env:OFFICEIMO_REQUIRE_PDF_RASTERIZER = $previousRequireRasterizer
    $env:OFFICEIMO_REQUIRE_PDF_RASTER_BASELINE_MATCH = $previousRequireRasterBaselineMatch
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
        $sourceArtifacts = @(
            foreach ($fileName in @($scenario.sourceReviewFiles)) {
                if ([string]::IsNullOrWhiteSpace($fileName)) {
                    continue
                }

                $artifactPath = Join-Path $resolvedOutputPath $fileName
                if (-not (Test-Path -LiteralPath $artifactPath)) {
                    throw "Manifest scenario '$($scenario.id)' expected source review artifact '$fileName', but it was not generated in $resolvedOutputPath."
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
            sourceArtifacts = $sourceArtifacts
        }
    }
)

$qualityContract = $scenarioManifest.qualityContract

$proofSummary = [pscustomobject]@{
    version = $scenarioManifest.version
    generatedAt = $generatedAt
    commit = $commit
    outputDirectory = $resolvedOutputPath
    manifest = 'conversion-scenarios.json'
    converterCatalog = @($scenarioManifest.converterCatalog)
    compositionRoutes = @($scenarioManifest.compositionRoutes)
    qualityContract = $qualityContract
    scenarios = $scenarioProof
}

$proofSummaryPath = Join-Path $resolvedOutputPath 'conversion-proof-summary.json'
$proofSummary | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $proofSummaryPath -Encoding UTF8

$supportLines = [System.Collections.Generic.List[string]]::new()
$supportLines.Add('# OfficeIMO PDF Conversion Support Matrix')
$supportLines.Add('')
$supportLines.Add('This matrix is generated from `Docs/pdf-conversion-scenarios.json`. Fidelity status describes the current evidence, not the intended destination.')
$supportLines.Add('')
$supportLines.Add("Premium claim rule: $($qualityContract.premiumClaimRule)")
$supportLines.Add('')
$supportLines.Add('| Source | Formats | Mode | Evidence status | Reference policy |')
$supportLines.Add('| --- | --- | --- | --- | --- |')
foreach ($converter in @($scenarioManifest.converterCatalog)) {
    $source = ([string]$converter.id).Replace('|', '\|')
    $formats = (@($converter.sourceFormats) -join ', ').Replace('|', '\|')
    $mode = ([string]$converter.conversionMode).Replace('|', '\|')
    $fidelityStatus = ([string]$converter.fidelityStatus).Replace('|', '\|')
    $referencePolicy = ([string]$converter.referencePolicy).Replace('|', '\|')
    $supportLines.Add("| $source | $formats | $mode | $fidelityStatus | $referencePolicy |")
}
$supportLines.Add('')
$supportLines.Add('## Capability Claims')
$supportLines.Add('')
$supportLines.Add('| Source | Capability | Fidelity level | Evidence scenarios |')
$supportLines.Add('| --- | --- | --- | --- |')
foreach ($converter in @($scenarioManifest.converterCatalog)) {
    if ($null -eq $converter.capabilityClaims) {
        continue
    }
    foreach ($claim in @($converter.capabilityClaims)) {
        $source = ([string]$converter.id).Replace('|', '\|')
        $capability = ([string]$claim.capability).Replace('|', '\|')
        $level = ([string]$claim.level).Replace('|', '\|')
        $evidence = (@($claim.evidenceScenarioIds) -join ', ').Replace('|', '\|')
        $supportLines.Add("| $source | $capability | $level | $evidence |")
    }
}
$supportLines.Add('')
$supportLines.Add('## Composed And Planned Routes')
$supportLines.Add('')
$supportLines.Add('| Route | Formats | Status | Diagnostic contract |')
$supportLines.Add('| --- | --- | --- | --- |')
foreach ($route in @($scenarioManifest.compositionRoutes)) {
    $routeId = ([string]$route.id).Replace('|', '\|')
    $formats = (@($route.sourceFormats) -join ', ').Replace('|', '\|')
    $routeStatus = ([string]$route.status).Replace('|', '\|')
    $diagnosticContract = ([string]$route.diagnosticContract).Replace('|', '\|')
    $supportLines.Add("| $routeId | $formats | $routeStatus | $diagnosticContract |")
}
[System.IO.File]::WriteAllLines($supportMatrixPath, $supportLines, [System.Text.Encoding]::UTF8)

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
$lines.Add("Generated support matrix: [pdf-conversion-support-matrix.md](pdf-conversion-support-matrix.md)")
$lines.Add('')
$lines.Add('## Premium Quality Contract')
$lines.Add('')
$lines.Add($qualityContract.goal)
$lines.Add('')
$lines.Add("Runtime ownership: $($qualityContract.runtimeOwnership)")
$lines.Add('')
$lines.Add('Required proof:')
$lines.Add('')
foreach ($proofItem in @($qualityContract.requiredProof)) {
    $lines.Add("- $proofItem")
}
$lines.Add('')
$lines.Add('Known limits:')
$lines.Add('')
foreach ($knownLimit in @($qualityContract.knownLimits)) {
    $lines.Add("- $knownLimit")
}
$lines.Add('')
$lines.Add('## Direct Converter Catalog')
$lines.Add('')
$lines.Add('| Source | Formats | Adapter | Mode | Evidence status |')
$lines.Add('| --- | --- | --- | --- | --- |')
foreach ($converter in @($scenarioManifest.converterCatalog)) {
    $source = ([string]$converter.id).Replace('|', '\|')
    $formats = (@($converter.sourceFormats) -join ', ').Replace('|', '\|')
    $adapter = ([string]$converter.adapter).Replace('|', '\|')
    $mode = ([string]$converter.conversionMode).Replace('|', '\|')
    $fidelityStatus = ([string]$converter.fidelityStatus).Replace('|', '\|')
    $lines.Add("| $source | $formats | $adapter | $mode | $fidelityStatus |")
}
$lines.Add('')
$lines.Add('## Composed And Planned Routes')
$lines.Add('')
$lines.Add('| Route | Formats | Status | Stages |')
$lines.Add('| --- | --- | --- | --- |')
foreach ($route in @($scenarioManifest.compositionRoutes)) {
    $routeId = ([string]$route.id).Replace('|', '\|')
    $formats = (@($route.sourceFormats) -join ', ').Replace('|', '\|')
    $routeStatus = ([string]$route.status).Replace('|', '\|')
    $stages = (@($route.stages) -join ' -> ').Replace('|', '\|')
    $lines.Add("| $routeId | $formats | $routeStatus | $stages |")
}
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
