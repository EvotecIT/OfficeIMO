param(
    [string] $SiteRoot = (Join-Path $PSScriptRoot '..\_site')
)

$ErrorActionPreference = 'Stop'

$resolvedSiteRoot = [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $SiteRoot))
$dataPath = Join-Path $resolvedSiteRoot 'data\benchmarks-excel.json'
$pagePath = Join-Path $resolvedSiteRoot 'benchmarks\index.html'
$scriptPath = Join-Path $resolvedSiteRoot 'js\benchmarks.js'

if (-not (Test-Path -LiteralPath $dataPath -PathType Leaf)) {
    throw "Benchmark data JSON was not published to '$dataPath'."
}

if (-not (Test-Path -LiteralPath $pagePath -PathType Leaf)) {
    throw "Benchmark page was not generated at '$pagePath'."
}

if (-not (Test-Path -LiteralPath $scriptPath -PathType Leaf)) {
    throw "Benchmark sort/filter script was not published to '$scriptPath'."
}

$data = Get-Content -LiteralPath $dataPath -Raw -Encoding UTF8 | ConvertFrom-Json
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

$pageHtml = Get-Content -LiteralPath $pagePath -Raw -Encoding UTF8
if ($pageHtml -notmatch 'data-excel-benchmarks' -or $pageHtml -notmatch 'data-benchmark-matrix') {
    throw "Benchmark page did not render the generated data dashboard."
}

if ($pageHtml -notmatch 'data-benchmark-family="excel"' -or
    $pageHtml -notmatch 'data-benchmark-family="csv"' -or
    $pageHtml -notmatch 'id="excel-matrix"' -or
    $pageHtml -notmatch 'Public baseline planned' -or
    $pageHtml -notmatch 'Word and PowerPoint') {
    throw "Benchmark page did not render the format-specific evidence hub and publication status."
}

if ($pageHtml -notmatch 'data-benchmark-sort="scenario"' -or $pageHtml -notmatch 'data-benchmark-filter="search"' -or $pageHtml -notmatch 'data-benchmark-reset' -or $pageHtml -notmatch 'data-benchmark-sort-mode' -or $pageHtml -notmatch '/js/benchmarks.js') {
    throw "Benchmark page did not render matrix sorting and filtering controls."
}

$scriptText = Get-Content -LiteralPath $scriptPath -Raw -Encoding UTF8
if ($scriptText -notmatch 'OfficeImoBenchmarkMatrix' -or $scriptText -notmatch 'sortBy' -or $scriptText -notmatch 'setFilter' -or $scriptText -notmatch 'setSortMetric' -or $scriptText -notmatch 'data-ratio-to-fastest') {
    throw "Benchmark sort/filter script does not expose the expected matrix behaviors."
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
