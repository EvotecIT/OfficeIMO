param(
    [string] $SiteRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
)

$ErrorActionPreference = 'Stop'
$failures = [System.Collections.Generic.List[string]]::new()

function Add-Failure([string] $Message) { $failures.Add($Message) }

$docsRoot = Join-Path $SiteRoot 'content\docs'
$tocPath = Join-Path $docsRoot 'toc.json'
$toc = Get-Content -LiteralPath $tocPath -Raw | ConvertFrom-Json
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
foreach ($card in @($showcase.cards)) {
    foreach ($requiredProperty in 'format', 'title', 'description', 'proof', 'source_url', 'guide_url', 'api_url') {
        if ([string]::IsNullOrWhiteSpace([string] $card.$requiredProperty)) {
            Add-Failure "Showcase card '$($card.title)' is missing '$requiredProperty'."
        }
    }
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
