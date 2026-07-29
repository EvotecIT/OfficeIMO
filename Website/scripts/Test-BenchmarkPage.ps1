param(
    [string] $SiteRoot = (Join-Path $PSScriptRoot '..\_site')
)

$ErrorActionPreference = 'Stop'

$resolvedSiteRoot = [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $SiteRoot))
$dataPath = Join-Path $resolvedSiteRoot 'data\benchmarks-excel.json'
$pagePath = Join-Path $resolvedSiteRoot 'benchmarks\index.html'
$scriptPath = Join-Path $resolvedSiteRoot 'js\benchmarks.js'
$tabularCatalogPath = Join-Path $resolvedSiteRoot 'data\benchmarks\tabular\index.json'
$stylePath = Join-Path $PSScriptRoot '..\themes\officeimo\assets\app.css'
$evidenceHelperPath = Join-Path $PSScriptRoot '..\..\Build\TabularBenchmarkEvidence.ps1'

. $evidenceHelperPath

if (-not (Test-Path -LiteralPath $dataPath -PathType Leaf)) {
    throw "Benchmark data JSON was not published to '$dataPath'."
}

if (-not (Test-Path -LiteralPath $pagePath -PathType Leaf)) {
    throw "Benchmark page was not generated at '$pagePath'."
}

if (-not (Test-Path -LiteralPath $scriptPath -PathType Leaf)) {
    throw "Benchmark sort/filter script was not published to '$scriptPath'."
}

if (-not (Test-Path -LiteralPath $tabularCatalogPath -PathType Leaf)) {
    throw "Tabular benchmark evidence catalog was not published to '$tabularCatalogPath'."
}

if (-not (Test-Path -LiteralPath $stylePath -PathType Leaf)) {
    throw "Benchmark layout styles were not found at '$stylePath'."
}

$contractEvidenceRoot = Join-Path ([System.IO.Path]::GetTempPath()) 'OfficeIMO-evidence-contract'
$net8Evidence = Get-TabularBenchmarkEvidenceLocation `
    -ComparisonId 'markpflug-65k-sales-v1-net8.0' `
    -Platform windows `
    -RunMode full `
    -StaticRoot $contractEvidenceRoot
$net10Evidence = Get-TabularBenchmarkEvidenceLocation `
    -ComparisonId 'markpflug-65k-sales-v1-net10.0' `
    -Platform windows `
    -RunMode full `
    -StaticRoot $contractEvidenceRoot
if ($net8Evidence.FileName -eq $net10Evidence.FileName -or
    $net8Evidence.ResultPath -eq $net10Evidence.ResultPath -or
    $net8Evidence.FileName -notmatch 'net8\.0' -or
    $net10Evidence.FileName -notmatch 'net10\.0') {
    throw 'Tabular benchmark payload paths do not keep framework comparison identities separate.'
}

$data = Get-Content -LiteralPath $dataPath -Raw -Encoding UTF8 | ConvertFrom-Json
$tabularCatalog = Get-Content -LiteralPath $tabularCatalogPath -Raw -Encoding UTF8 | ConvertFrom-Json
$rowCount = @($data.rows).Count
$summaryCount = @($data.summary).Count
$matrixRowCount = @($data.matrix.rows).Count

if ($rowCount -lt 1) {
    throw "Benchmark data JSON does not contain measurement rows."
}

if ($matrixRowCount -lt 1) {
    throw "Benchmark data JSON does not contain matrix rows."
}

if ($summaryCount -lt 1) {
    throw "Benchmark data JSON does not contain summary rows."
}

if ($tabularCatalog.schemaVersion -ne 2 -or
    @($tabularCatalog.expectedPlatforms).Count -ne 3 -or
    -not (@($tabularCatalog.availability).Where({ $_.platform -eq 'macos' -and -not $_.available }))) {
    throw 'Tabular benchmark catalog does not expose the versioned three-platform model and explicit missing macOS lane.'
}

$pageHtml = Get-Content -LiteralPath $pagePath -Raw -Encoding UTF8
if ($pageHtml -notmatch 'data-excel-benchmarks' -or $pageHtml -notmatch 'data-benchmark-matrix') {
    throw "Benchmark page did not render the generated data dashboard."
}

if ($pageHtml -notmatch 'data-tabular-benchmarks' -or
    $pageHtml -notmatch 'data-comparison-id="markpflug-65k-sales-v1-net10\.0"' -or
    $pageHtml -notmatch 'data-tabular-platform="windows"' -or
    $pageHtml -notmatch 'data-tabular-platform="linux"' -or
    $pageHtml -notmatch 'data-tabular-platform="macos"' -or
    $pageHtml -notmatch 'data-tabular-mode="full"' -or
    $pageHtml -notmatch 'data-tabular-mode="quick"') {
    throw 'Benchmark page did not render the tabular platform and evidence selectors.'
}

if ($pageHtml -notmatch 'data-benchmark-family="excel"' -or
    $pageHtml -notmatch 'data-benchmark-family="csv"' -or
    $pageHtml -notmatch 'id="excel-matrix"' -or
    $pageHtml -notmatch 'Coverage boundary' -or
    $pageHtml -notmatch 'Word and PowerPoint' -or
    $pageHtml -notmatch '/docs/capabilities/benchmarks/') {
    throw "Benchmark page did not render the format-specific evidence hub, coverage boundary, and reproduction guide."
}

if ($pageHtml -match 'github\.com/EvotecIT/OfficeIMO/(?:blob|tree)/main') {
    throw "Benchmark page contains a public evidence link to the nonexistent OfficeIMO 'main' branch."
}

if ($pageHtml -notmatch 'data-benchmark-sort="scenario"' -or $pageHtml -notmatch 'data-benchmark-filter="search"' -or $pageHtml -notmatch 'data-benchmark-reset' -or $pageHtml -notmatch 'data-benchmark-sort-mode' -or $pageHtml -notmatch '/js/benchmarks.js') {
    throw "Benchmark page did not render matrix sorting and filtering controls."
}

$scriptText = Get-Content -LiteralPath $scriptPath -Raw -Encoding UTF8
if ($scriptText -notmatch 'OfficeImoBenchmarkMatrix' -or $scriptText -notmatch 'sortBy' -or $scriptText -notmatch 'setFilter' -or $scriptText -notmatch 'setSortMetric' -or $scriptText -notmatch 'data-ratio-to-fastest') {
    throw "Benchmark sort/filter script does not expose the expected matrix behaviors."
}

if ($scriptText -notmatch 'benchmark-os' -or
    $scriptText -notmatch 'benchmark-mode' -or
    $scriptText -notmatch 'candidate\.comparisonId === selectedComparison' -or
    $scriptText -notmatch 'candidate\.publish === true' -or
    $scriptText -notmatch 'compatibilityIssues' -or
    $pageHtml -notmatch 'Quick results are diagnostic only') {
    throw 'Tabular benchmark selector does not preserve shareable platform/mode state and evidence safety labels.'
}

$styleText = Get-Content -LiteralPath $stylePath -Raw -Encoding UTF8
if ($styleText -notmatch '\.imo-benchmark-hub\{[^}]*grid-template-columns:minmax\(0,1fr\)') {
    throw "Benchmark hub does not constrain wide children to a responsive grid track."
}

if ($styleText -notmatch '\.imo-benchmark-explorer\{[^}]*justify-self:center[^}]*width:min\(1600px,calc\(100vw - 3rem\)\)') {
    throw "Benchmark explorer does not own the centered wide layout."
}

if ($styleText -notmatch '\.imo-benchmark-dashboard\{width:100%;margin:0 0 3rem\}' -or
    $styleText -match '\.imo-benchmark-dashboard\{[^}]*100vw') {
    throw "Benchmark dashboard can overflow and be clipped by its explorer."
}

if ($pageHtml -match 'Loading benchmark data') {
    throw "Benchmark page still depends on client-side data loading."
}

$renderedMatrixRows = ([regex]::Matches($pageHtml, '<tr[^>]*data-benchmark-row[^>]*>\s*<td class="imo-benchmark-scenario"[^>]*data-label="Scenario"')).Count
if ($renderedMatrixRows -lt $matrixRowCount) {
    throw "Benchmark page rendered $renderedMatrixRows matrix rows, expected at least $matrixRowCount."
}

$sortableCells = ([regex]::Matches($pageHtml, 'data-library="[^"]+"')).Count
if ($sortableCells -lt $rowCount) {
    throw "Benchmark page rendered $sortableCells sortable library cells, expected at least $rowCount."
}

$responsiveCells = ([regex]::Matches($pageHtml, 'data-label="[^"]+"')).Count
if ($responsiveCells -lt $sortableCells) {
    throw "Benchmark page did not render responsive data labels for matrix cells."
}

$ratioSortCells = ([regex]::Matches($pageHtml, 'data-ratio-to-fastest="[^"]+"')).Count
if ($ratioSortCells -lt $rowCount) {
    throw "Benchmark page did not render ratio sort metadata for measured cells."
}

if ($pageHtml -match 'Strongest OfficeIMO Wins' -or $pageHtml -match 'Optimization Targets') {
    throw "Benchmark page still contains the old win/loss commentary panels."
}

Write-Host "Benchmark page verified: $matrixRowCount matrix rows, $rowCount measurement rows, $summaryCount summary rows."
