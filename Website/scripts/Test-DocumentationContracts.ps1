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
    if ($raw -match '(?m)^#\s+') {
        Add-Failure "'$([System.IO.Path]::GetRelativePath($docsRoot, $doc.FullName))' contains a body H1; the docs layout already renders the page title."
    }
    if ($raw -match '/examples/pswriteoffice/') {
        Add-Failure "'$([System.IO.Path]::GetRelativePath($docsRoot, $doc.FullName))' links to the retired /examples/pswriteoffice route."
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
    projectCount = 145
    productionComponentCount = 89
    testProjectCount = 29
    benchmarkProjectCount = 12
    validationProjectCount = 16
    apiReferenceCount = 17
    conceptualPageCount = 69
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
