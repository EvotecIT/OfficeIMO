param(
    [string] $SiteRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
)

$ErrorActionPreference = 'Stop'
$failures = [System.Collections.Generic.List[string]]::new()

function Add-Failure([string] $Message) { $failures.Add($Message) }

function Test-ResponsiveImageRule([string] $Path, [string] $Label) {
    $css = Get-Content -LiteralPath $Path -Raw
    if ($css -notmatch '(?s)(?:^|})\s*img\s*\{[^}]*\bheight\s*:\s*auto\s*;?[^}]*\}') {
        Add-Failure "$Label must preserve intrinsic image aspect ratios with an img { height: auto; } rule."
    }
}

function Get-PngDimensions([string] $Path) {
    [byte[]] $bytes = [System.IO.File]::ReadAllBytes($Path)
    if ($bytes.Length -lt 24 -or
        $bytes[0] -ne 137 -or $bytes[1] -ne 80 -or
        $bytes[2] -ne 78 -or $bytes[3] -ne 71) {
        throw "'$Path' is not a PNG file with an IHDR chunk."
    }

    $width = ([int] $bytes[16] -shl 24) -bor
        ([int] $bytes[17] -shl 16) -bor
        ([int] $bytes[18] -shl 8) -bor
        [int] $bytes[19]
    $height = ([int] $bytes[20] -shl 24) -bor
        ([int] $bytes[21] -shl 16) -bor
        ([int] $bytes[22] -shl 8) -bor
        [int] $bytes[23]
    return [pscustomobject] @{ Width = $width; Height = $height }
}

$docsRoot = Join-Path $SiteRoot 'content\docs'
$tocPath = Join-Path $docsRoot 'toc.json'
$toc = Get-Content -LiteralPath $tocPath -Raw | ConvertFrom-Json
$criticalCssPath = Join-Path $SiteRoot 'themes\officeimo\critical.css'
$appCssPath = Join-Path $SiteRoot 'themes\officeimo\assets\app.css'
Test-ResponsiveImageRule -Path $criticalCssPath -Label 'Critical site CSS'
Test-ResponsiveImageRule -Path $appCssPath -Label 'Full site CSS'

$tocEntries = @($toc | ForEach-Object {
    if ($_.path) { $_ }
    @($_.items)
}) | Where-Object { $_.path }

foreach ($entry in $tocEntries) {
    $sourcePath = Join-Path $docsRoot ([string] $entry.path)
    if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
        Add-Failure "Navigation entry '$($entry.title)' points to missing source '$($entry.path)'."
    }
}

$docs = @(Get-ChildItem -LiteralPath $docsRoot -Recurse -File -Filter '*.md')
foreach ($doc in $docs) {
    $raw = Get-Content -LiteralPath $doc.FullName -Raw
    $relativePath = [System.IO.Path]::GetRelativePath($docsRoot, $doc.FullName)
    if ($raw -match '(?m)^#\s+') {
        Add-Failure "'$relativePath' contains a body H1; the docs layout already renders the page title."
    }
    if ($raw -match '/examples/pswriteoffice/') {
        Add-Failure "'$relativePath' links to the retired /examples/pswriteoffice route."
    }
    if ($raw -match '(?i)Market Position|Out Of Scope Here|should position itself|What we do not have yet') {
        Add-Failure "'$relativePath' contains internal positioning or planning copy that does not belong in customer documentation."
    }
}

$publicContentFiles = @(
    Get-ChildItem -LiteralPath (Join-Path $SiteRoot 'content') -Recurse -File |
        Where-Object Extension -In '.md', '.html'
    Get-ChildItem -LiteralPath (Join-Path $SiteRoot 'data') -File |
        Where-Object Extension -EQ '.json'
)
foreach ($publicContentFile in $publicContentFiles) {
    $publicContent = Get-Content -LiteralPath $publicContentFile.FullName -Raw
    if ($publicContent -match 'github\.com/EvotecIT/OfficeIMO/(?:blob|tree)/main(?:/|")') {
        Add-Failure "'$([System.IO.Path]::GetRelativePath($SiteRoot, $publicContentFile.FullName))' uses the nonexistent OfficeIMO 'main' branch in a public link."
    }
}

$excelProductPath = Join-Path $SiteRoot 'content\products\excel.md'
$excelProduct = Get-Content -LiteralPath $excelProductPath -Raw
if ($excelProduct -notmatch 'workbook\.AddWorksheet\("Q4 Sales"\)') {
    Add-Failure 'The OfficeIMO.Excel product quick start must use ExcelDocument.AddWorksheet.'
}
if ($excelProduct -notmatch '(?s)sheet\.AddTable\(\s*\$"A1:C\{totalsRow\}",\s*hasHeader:\s*true,\s*name:\s*"SalesTable",\s*style:\s*TableStyle\.TableStyleMedium9\)') {
    Add-Failure 'The OfficeIMO.Excel product quick start must use the supported ExcelSheet.AddTable signature and TableStyle enum.'
}
if ($excelProduct -notmatch 'sheet\.SetTableTotalsByName\(') {
    Add-Failure 'The OfficeIMO.Excel product quick start must use the supported named-table totals API.'
}

$powerPointImageExportPath = Join-Path $docsRoot 'powerpoint\image-export\index.md'
$powerPointImageExport = Get-Content -LiteralPath $powerPointImageExportPath -Raw
if ($powerPointImageExport -notmatch 'PowerPointPresentation\.Load\("Quarterly-Review\.pptx"\)') {
    Add-Failure 'The PowerPoint image-export guide must load existing presentations through PowerPointPresentation.Load.'
}

$showcasePath = Join-Path $SiteRoot 'data\showcase.json'
$showcase = Get-Content -LiteralPath $showcasePath -Raw | ConvertFrom-Json
if (@($showcase.cards).Count -lt 6) {
    Add-Failure 'The showcase must retain at least six evidence-backed workflows.'
}
foreach ($card in @($showcase.cards)) {
    foreach ($requiredProperty in @(
        'format', 'title', 'description', 'preview_kind', 'proof', 'limit',
        'artifact_url', 'artifact_label', 'evidence_url', 'evidence_label',
        'source_url', 'guide_url', 'api_url'
    )) {
        if ([string]::IsNullOrWhiteSpace([string] $card.$requiredProperty)) {
            Add-Failure "Showcase card '$($card.title)' is missing '$requiredProperty'."
        }
    }

    foreach ($localUrlProperty in 'artifact_url', 'evidence_url') {
        $localUrl = [string] $card.$localUrlProperty
        if (-not $localUrl.StartsWith('/downloads/showcase/', [StringComparison]::Ordinal)) {
            Add-Failure "Showcase card '$($card.title)' uses a non-evidence '$localUrlProperty' URL: $localUrl"
            continue
        }

        $localPath = Join-Path (Join-Path $SiteRoot 'static') $localUrl.TrimStart([char[]] '/')
        if (-not (Test-Path -LiteralPath $localPath -PathType Leaf)) {
            Add-Failure "Showcase card '$($card.title)' points to missing evidence '$localUrl'."
        }
    }

    if ($card.preview_kind -eq 'image') {
        if ([string]::IsNullOrWhiteSpace([string] $card.image) -or
            [string] $card.image -match '/images/formats/') {
            Add-Failure "Showcase card '$($card.title)' must use generated visual evidence, not a format icon."
        } else {
            $previewPath = Join-Path (Join-Path $SiteRoot 'static') ([string] $card.image).TrimStart([char[]] '/')
            if (-not (Test-Path -LiteralPath $previewPath -PathType Leaf)) {
                Add-Failure "Showcase card '$($card.title)' points to missing preview '$($card.image)'."
            } elseif ([System.IO.Path]::GetExtension($previewPath) -ieq '.png') {
                $dimensions = Get-PngDimensions $previewPath
                if ($dimensions.Width -ne [int] $card.image_width -or
                    $dimensions.Height -ne [int] $card.image_height) {
                    Add-Failure "Showcase card '$($card.title)' declares $($card.image_width)x$($card.image_height), but its PNG is $($dimensions.Width)x$($dimensions.Height)."
                }
            }
        }
    } elseif ($card.preview_kind -eq 'reader') {
        if (@($card.preview_items).Count -lt 3 -or [string]::IsNullOrWhiteSpace([string] $card.preview_title)) {
            Add-Failure "Showcase card '$($card.title)' has an incomplete Reader preview."
        }
    } elseif ($card.preview_kind -eq 'onenote') {
        foreach ($previewProperty in 'preview_title', 'preview_heading', 'preview_body', 'preview_item') {
            if ([string]::IsNullOrWhiteSpace([string] $card.$previewProperty)) {
                Add-Failure "Showcase card '$($card.title)' is missing OneNote preview field '$previewProperty'."
            }
        }
    } else {
        Add-Failure "Showcase card '$($card.title)' has unsupported preview kind '$($card.preview_kind)'."
    }
}

$showcaseManifestPath = Join-Path $SiteRoot 'static\downloads\showcase\manifest.json'
$showcaseManifest = Get-Content -LiteralPath $showcaseManifestPath -Raw | ConvertFrom-Json
if ($showcaseManifest.schema -ne 'officeimo.showcase-evidence' -or $showcaseManifest.schemaVersion -ne 1) {
    Add-Failure 'The showcase evidence manifest has an unsupported schema.'
}
$manifestPaths = @($showcaseManifest.artifacts | ForEach-Object { [string] $_.path })
foreach ($artifact in @($showcaseManifest.artifacts)) {
    $artifactPath = Join-Path (Join-Path $SiteRoot 'static') ([string] $artifact.path).TrimStart([char[]] '/')
    if (-not (Test-Path -LiteralPath $artifactPath -PathType Leaf)) {
        Add-Failure "Showcase manifest artifact '$($artifact.id)' is missing '$($artifact.path)'."
        continue
    }

    $file = Get-Item -LiteralPath $artifactPath
    if ($file.Length -ne [long] $artifact.bytes) {
        Add-Failure "Showcase manifest artifact '$($artifact.id)' has a stale byte count."
    }
    $actualHash = (Get-FileHash -LiteralPath $artifactPath -Algorithm SHA256).Hash.ToLowerInvariant()
    if ($actualHash -ne [string] $artifact.sha256) {
        Add-Failure "Showcase manifest artifact '$($artifact.id)' has a stale SHA-256 hash."
    }
}
foreach ($card in @($showcase.cards)) {
    foreach ($evidenceUrl in @([string] $card.artifact_url, [string] $card.evidence_url)) {
        if ($manifestPaths -notcontains $evidenceUrl) {
            Add-Failure "Showcase card '$($card.title)' evidence '$evidenceUrl' is missing from the manifest."
        }
    }
}

$readerCard = @($showcase.cards | Where-Object preview_kind -eq 'reader') | Select-Object -First 1
$readerEvidencePath = Join-Path (Join-Path $SiteRoot 'static') ([string] $readerCard.artifact_url).TrimStart([char[]] '/')
$readerEvidenceRaw = Get-Content -LiteralPath $readerEvidencePath -Raw
$readerEvidence = $readerEvidenceRaw | ConvertFrom-Json
if ($readerEvidence.schemaId -ne 'officeimo.document.read-result' -or
    $readerEvidence.schemaVersion -ne 6 -or
    $readerEvidence.kind -ne 'PowerPoint' -or
    @($readerEvidence.chunks).Count -ne 4) {
    Add-Failure 'The bundled Reader proof is not the expected four-slide schema-v6 PowerPoint result.'
}
if ($readerEvidenceRaw -match '"path"\s*:\s*"[A-Za-z]:\\\\') {
    Add-Failure 'The bundled Reader proof leaks a machine-local repository path.'
}
foreach ($previewItem in @($readerCard.preview_items)) {
    if ([string] $readerEvidence.markdown -notmatch [regex]::Escape([string] $previewItem)) {
        Add-Failure "Reader preview item '$previewItem' is not present in the bundled Reader output."
    }
}

$oneNoteCard = @($showcase.cards | Where-Object preview_kind -eq 'onenote') | Select-Object -First 1
$oneNoteHtmlPath = Join-Path $SiteRoot 'static\downloads\showcase\onenote\offline-planning.html.txt'
$oneNoteHtml = Get-Content -LiteralPath $oneNoteHtmlPath -Raw
$oneNoteText = $oneNoteHtml -replace '<[^>]+>', ''
foreach ($previewProperty in 'preview_title', 'preview_heading', 'preview_body', 'preview_item') {
    if ($oneNoteText -notmatch [regex]::Escape([string] $oneNoteCard.$previewProperty)) {
        Add-Failure "OneNote preview field '$previewProperty' is not present in the bundled HTML export."
    }
}

$downloadsCatalogPath = Join-Path $SiteRoot 'data\downloads_catalog.json'
$downloadsCatalog = Get-Content -LiteralPath $downloadsCatalogPath -Raw | ConvertFrom-Json
$downloadPackages = @($downloadsCatalog.sections.packages)
$rtfDownload = @($downloadPackages | Where-Object registry_package -eq 'OfficeIMO.Rtf') | Select-Object -First 1
if ($null -eq $rtfDownload -or
    $rtfDownload.product_url -ne '/products/rtf/' -or
    $rtfDownload.docs_url -ne '/docs/rtf/') {
    Add-Failure 'The Downloads catalog must expose OfficeIMO.Rtf with its dedicated product and documentation routes.'
}
$epubDownload = @($downloadPackages | Where-Object registry_package -eq 'OfficeIMO.Epub') | Select-Object -First 1
if ($null -eq $epubDownload -or
    $epubDownload.product_url -ne '/products/epub/' -or
    $epubDownload.docs_url -ne '/docs/epub/') {
    Add-Failure 'The Downloads catalog must route OfficeIMO.Epub to its dedicated product and documentation pages.'
}

$openTextFormatsPath = Join-Path $docsRoot 'pswriteoffice\open-text-formats\index.md'
$openTextFormats = Get-Content -LiteralPath $openTextFormatsPath -Raw
if ($openTextFormats -notmatch '(?m)^\s*-\s+/docs/pswriteoffice/markdown/\s*$') {
    Add-Failure 'The retired PSWriteOffice Markdown URL is not preserved as an alias of the open and text formats guide.'
}

$catalogPath = Join-Path $SiteRoot 'data\documentation_catalog.json'
$catalog = Get-Content -LiteralPath $catalogPath -Raw | ConvertFrom-Json
if ($catalog.repository.productionComponentCount -ne @($catalog.components).Count) {
    Add-Failure 'The OfficeIMO component summary does not match the generated component list.'
}
$expectedRepositoryCounts = [ordered]@{
    projectCount = 146
    productionComponentCount = 89
    testProjectCount = 29
    benchmarkProjectCount = 12
    validationProjectCount = 17
    apiReferenceCount = 17
    conceptualPageCount = 89
}
foreach ($expectedCount in $expectedRepositoryCounts.GetEnumerator()) {
    $actual = [int] $catalog.repository.($expectedCount.Key)
    if ($actual -ne $expectedCount.Value) {
        Add-Failure "The OfficeIMO $($expectedCount.Key) is $actual; expected $($expectedCount.Value) on every operating system."
    }
}
if (@($catalog.components | Where-Object { [string]::IsNullOrWhiteSpace($_.description) }).Count -gt 0) {
    Add-Failure 'One or more OfficeIMO catalog components have no description.'
}
$rtfCatalogComponents = @($catalog.components | Where-Object name -Like 'OfficeIMO.Rtf*')
if ($rtfCatalogComponents.Count -eq 0 -or
    @($rtfCatalogComponents | Where-Object docsUrl -ne '/docs/rtf/').Count -gt 0) {
    Add-Failure 'Every generated OfficeIMO.Rtf catalog entry must route to the dedicated RTF guide.'
}
$epubCatalogComponents = @($catalog.components | Where-Object name -Like 'OfficeIMO.Epub*')
if ($epubCatalogComponents.Count -eq 0 -or
    @($epubCatalogComponents | Where-Object docsUrl -ne '/docs/epub/').Count -gt 0) {
    Add-Failure 'Every generated OfficeIMO.Epub catalog entry must route to the dedicated EPUB guide.'
}

$aotMatrixPath = Join-Path $SiteRoot 'static\data\aot-compatibility.json'
$aotMatrix = Get-Content -LiteralPath $aotMatrixPath -Raw | ConvertFrom-Json
if ($aotMatrix.summary.productionProjectCount -ne $catalog.repository.productionComponentCount) {
    Add-Failure 'The NativeAOT matrix does not account for every production project.'
}
if ($aotMatrix.summary.nativeAotValidatedProjectCount -ne 88) {
    Add-Failure "The NativeAOT matrix validates $($aotMatrix.summary.nativeAotValidatedProjectCount) projects; expected 88."
}
if ($aotMatrix.summary.fullyRootedLibraryCount -ne 85 -or
    $aotMatrix.summary.boundedWorkflowLibraryCount -ne 1 -or
    $aotMatrix.summary.nativeExecutableCount -ne 2 -or
    $aotMatrix.summary.managedWindowsProjectCount -ne 1) {
    Add-Failure 'The NativeAOT classification totals changed without updating the customer-facing contract.'
}
if (@($aotMatrix.components).Count -ne $catalog.repository.productionComponentCount) {
    Add-Failure 'The NativeAOT component list is incomplete.'
}

$powerShellCatalogPath = Join-Path $SiteRoot 'data\pswriteoffice_command_catalog.json'
$powerShellCatalog = Get-Content -LiteralPath $powerShellCatalogPath -Raw | ConvertFrom-Json
if ($powerShellCatalog.module.commandCount -ne 464) {
    Add-Failure "The PSWriteOffice snapshot has $($powerShellCatalog.module.commandCount) commands; expected the authoritative 464-command surface."
}
if ((@($powerShellCatalog.families | Measure-Object commandCount -Sum).Sum) -ne $powerShellCatalog.module.commandCount) {
    Add-Failure 'The PSWriteOffice family totals do not cover each command exactly once.'
}
if ($powerShellCatalog.module.aliasCount -ne 354) {
    Add-Failure "The PSWriteOffice snapshot has $($powerShellCatalog.module.aliasCount) aliases; expected 354."
}

if ($failures.Count -gt 0) {
    throw "Documentation contract validation failed:`n - $($failures -join "`n - ")"
}

[PSCustomObject]@{
    DocumentationPageCount = $docs.Count
    NavigationEntryCount = $tocEntries.Count
    ProductionComponentCount = $catalog.repository.productionComponentCount
    PowerShellCommandCount = $powerShellCatalog.module.commandCount
    Status = 'passed'
}
