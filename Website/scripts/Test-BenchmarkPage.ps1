param(
    [string] $SiteRoot = (Join-Path $PSScriptRoot '..\_site')
)

$ErrorActionPreference = 'Stop'

$resolvedSiteRoot = if ([System.IO.Path]::IsPathRooted($SiteRoot)) {
    [System.IO.Path]::GetFullPath($SiteRoot)
} else {
    [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $SiteRoot))
}
$dataPath = Join-Path $resolvedSiteRoot 'data\benchmarks-excel.json'
$pagePath = Join-Path $resolvedSiteRoot 'benchmarks\index.html'
$scriptPath = Join-Path $resolvedSiteRoot 'js\benchmarks.js'
$comparisonCatalogPath = Join-Path $resolvedSiteRoot 'data\benchmarks\library-comparisons\index.json'
$stylePath = Join-Path $PSScriptRoot '..\themes\officeimo\assets\app.css'
$evidenceHelperPath = Join-Path $PSScriptRoot '..\..\Build\BenchmarkEvidence.ps1'
$comparisonRunnerPath = Join-Path $PSScriptRoot '..\..\Build\Run-LibraryComparisonBenchmarks.ps1'

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

if (-not (Test-Path -LiteralPath $comparisonCatalogPath -PathType Leaf)) {
    throw "Library comparison evidence catalog was not published to '$comparisonCatalogPath'."
}

if (-not (Test-Path -LiteralPath $stylePath -PathType Leaf)) {
    throw "Benchmark layout styles were not found at '$stylePath'."
}

if (-not (Test-Path -LiteralPath $comparisonRunnerPath -PathType Leaf)) {
    throw "Library comparison runner was not found at '$comparisonRunnerPath'."
}

$contractEvidenceRoot = Join-Path ([System.IO.Path]::GetTempPath()) 'OfficeIMO-evidence-contract'
$net8Evidence = Get-BenchmarkEvidenceLocation `
    -ComparisonId 'markpflug-65k-csv-decoded-net8.0' `
    -Platform windows `
    -RunMode full `
    -StaticRoot $contractEvidenceRoot
$net10Evidence = Get-BenchmarkEvidenceLocation `
    -ComparisonId 'markpflug-65k-csv-decoded-net10.0' `
    -Platform windows `
    -RunMode full `
    -StaticRoot $contractEvidenceRoot
if ($net8Evidence.FileName -eq $net10Evidence.FileName -or
    $net8Evidence.ResultPath -eq $net10Evidence.ResultPath -or
    $net8Evidence.FileName -notmatch 'net8\.0' -or
    $net10Evidence.FileName -notmatch 'net10\.0') {
    throw 'Library comparison payload paths do not keep framework identities separate.'
}

$data = Get-Content -LiteralPath $dataPath -Raw -Encoding UTF8 | ConvertFrom-Json
$comparisonCatalog = Get-Content -LiteralPath $comparisonCatalogPath -Raw -Encoding UTF8 | ConvertFrom-Json
$rowCount = @($data.rows).Count
$summaryCount = @($data.summary).Count
$matrixRowCount = @($data.matrix.rows).Count

if ($data.schemaVersion -ne 2 -or
    $data.platform -ne 'unrecorded' -or
    $data.runMode -ne 'unrecorded' -or
    $data.publish) {
    throw 'Legacy Excel matrix data does not explicitly preserve its unrecorded OS and run mode.'
}

$comparisonRunnerText = Get-Content -LiteralPath $comparisonRunnerPath -Raw -Encoding UTF8
if ($comparisonRunnerText -notmatch 'if \(\$catalogEligible -and \$gitDirty\)' -or
    $comparisonRunnerText -notmatch 'Cataloged benchmark evidence requires a clean Git worktree') {
    throw 'Catalog-eligible benchmark runs do not enforce source-commit provenance from a clean worktree.'
}

if ($rowCount -lt 1) {
    throw "Benchmark data JSON does not contain measurement rows."
}

if ($matrixRowCount -lt 1) {
    throw "Benchmark data JSON does not contain matrix rows."
}

if ($summaryCount -lt 1) {
    throw "Benchmark data JSON does not contain summary rows."
}

if ($comparisonCatalog.schemaVersion -ne 3 -or
    @($comparisonCatalog.expectedPlatforms).Count -ne 3 -or
    -not (@($comparisonCatalog.availability).Where({ $_.platform -eq 'macos' -and -not $_.available }))) {
    throw 'Library comparison catalog does not expose the versioned three-platform model and explicit missing macOS lane.'
}

foreach ($entry in @($comparisonCatalog.entries)) {
    if ([string]::IsNullOrWhiteSpace([string] $entry.resultPath) -or
        [string]::IsNullOrWhiteSpace([string] $entry.resultSha256)) {
        throw "Library comparison catalog entry '$($entry.comparisonId)' is missing result-path integrity metadata."
    }

    $relativeResultPath = ([string] $entry.resultPath).TrimStart('/').Replace(
        '/',
        [System.IO.Path]::DirectorySeparatorChar)
    $publishedResultPath = Join-Path $resolvedSiteRoot $relativeResultPath
    if (-not (Test-Path -LiteralPath $publishedResultPath -PathType Leaf)) {
        throw "Library comparison result '$($entry.resultPath)' was not published."
    }

    $actualResultSha256 = (Get-FileHash -LiteralPath $publishedResultPath -Algorithm SHA256).Hash
    if (-not [string]::Equals(
            $actualResultSha256,
            [string] $entry.resultSha256,
            [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Library comparison result '$($entry.resultPath)' does not match its catalog SHA-256."
    }

    $sourceCommit = [string] $entry.compatibility.'benchmark.workload.sourceCommit'
    $gitSha = [string] $entry.compatibility.gitSha
    if ([string]::IsNullOrWhiteSpace($sourceCommit) -or
        [string]::IsNullOrWhiteSpace($gitSha) -or
        -not [string]::Equals($sourceCommit, $gitSha, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Library comparison result '$($entry.resultPath)' does not bind sourceCommit and gitSha to the same measured commit."
    }
}

$readComparisonIds = @(
    'markpflug-65k-csv-decoded-net10.0',
    'markpflug-65k-xlsx-typed-net10.0',
    'markpflug-65k-xlsb-typed-net10.0'
)
foreach ($comparisonId in $readComparisonIds) {
    $windowsModes = @(
        $comparisonCatalog.entries |
            Where-Object {
                $_.comparisonId -eq $comparisonId -and
                $_.platform -eq 'windows'
            } |
            ForEach-Object runMode |
            Sort-Object -Unique
    )
    if ('full' -notin $windowsModes -or 'quick' -notin $windowsModes) {
        throw "Library comparison '$comparisonId' does not publish both full and quick Windows evidence."
    }
}


$writeComparisonScenarios = [ordered]@{
    'csv-25k-datareader-write-net10.0' = @('OfficeIMO_WriteDataReader', 'Sylvan_WriteDataReader')
    'xlsx-25k-datareader-write-net10.0' = @('OfficeIMO', 'SpreadCheetah', 'Sylvan', 'LargeXlsx')
}
foreach ($writeComparison in $writeComparisonScenarios.GetEnumerator()) {
    $writeEntries = @($comparisonCatalog.entries | Where-Object { $_.comparisonId -eq $writeComparison.Key })
    $macQuickEntry = $writeEntries | Where-Object { $_.platform -eq 'macos' -and $_.runMode -eq 'quick' } | Select-Object -First 1
    $macQuickAvailability = $comparisonCatalog.availability | Where-Object {
        $_.comparisonId -eq $writeComparison.Key -and
        $_.runMode -eq 'quick' -and
        $_.platform -eq 'macos' -and
        $_.available
    } | Select-Object -First 1
    if (-not $macQuickEntry -or -not $macQuickAvailability) {
        throw "Validated write evidence '$($writeComparison.Key)' is missing from the platform-aware catalog."
    }

    $relativeWriteResultPath = ([string] $macQuickEntry.resultPath).TrimStart('/').Replace(
        '/',
        [System.IO.Path]::DirectorySeparatorChar)
    $writeResult = Get-Content -LiteralPath (Join-Path $resolvedSiteRoot $relativeWriteResultPath) -Raw -Encoding UTF8 | ConvertFrom-Json
    $actualScenarios = @($writeResult.summary | ForEach-Object scenario | Sort-Object -Unique)
    $missingScenarios = @($writeComparison.Value | Where-Object { $_ -notin $actualScenarios })
    if ($missingScenarios.Count -gt 0) {
        throw "Write evidence '$($writeComparison.Key)' is missing scenarios: $($missingScenarios -join ', ')."
    }
}

$pageHtml = Get-Content -LiteralPath $pagePath -Raw -Encoding UTF8
if ($pageHtml -notmatch 'data-excel-benchmarks' -or $pageHtml -notmatch 'data-benchmark-matrix') {
    throw "Benchmark page did not render the generated data dashboard."
}

if ($pageHtml -notmatch 'data-library-comparison-benchmarks' -or
    $pageHtml -notmatch 'data-comparison-id="markpflug-65k-csv-decoded-net10\.0"' -or
    $pageHtml -notmatch 'data-library-comparison-workload="csv-25k-datareader-write-net10\.0"' -or
    $pageHtml -notmatch 'data-library-comparison-workload="xlsx-25k-datareader-write-net10\.0"' -or
    $pageHtml -notmatch 'data-library-comparison-workload="markpflug-65k-xlsx-typed-net10\.0"' -or
    $pageHtml -notmatch 'data-library-comparison-workload="markpflug-65k-xlsb-typed-net10\.0"' -or
    $pageHtml -notmatch 'data-library-comparison-platform="windows"' -or
    $pageHtml -notmatch 'data-library-comparison-platform="linux"' -or
    $pageHtml -notmatch 'data-library-comparison-platform="macos"' -or
    $pageHtml -notmatch 'data-library-comparison-mode="full"' -or
    $pageHtml -notmatch 'data-library-comparison-mode="quick"') {
    throw 'Benchmark page did not render the workload, platform, and evidence selectors.'
}

if ($pageHtml -notmatch 'data-benchmark-family="excel"' -or
    $pageHtml -notmatch 'data-benchmark-family="csv"' -or
    $pageHtml -notmatch 'id="excel-matrix"' -or
    $pageHtml -notmatch 'id="historical-benchmark-evidence"' -or
    $pageHtml -notmatch 'OS not recorded' -or
    $pageHtml -notmatch 'Run mode not recorded' -or
    $pageHtml -notmatch 'Coverage boundary' -or
    $pageHtml -notmatch 'Word and PowerPoint' -or
    $pageHtml -notmatch '/docs/capabilities/benchmarks/') {
    throw "Benchmark page did not render the format-specific evidence hub, coverage boundary, and reproduction guide."
}

if ($pageHtml -match 'github\.com/EvotecIT/OfficeIMO/(?:blob|tree)/main') {
    throw "Benchmark page contains a public evidence link to the nonexistent OfficeIMO 'main' branch."
}

if ($pageHtml -notmatch 'data-benchmark-sort="scenario"' -or
    $pageHtml -notmatch 'data-benchmark-filter="search"' -or
    $pageHtml -notmatch 'data-benchmark-filter="platform"' -or
    $pageHtml -notmatch 'data-benchmark-filter="runMode"' -or
    $pageHtml -notmatch 'data-platform="unrecorded"' -or
    $pageHtml -notmatch 'data-run-mode="unrecorded"' -or
    $pageHtml -notmatch 'data-benchmark-empty' -or
    $pageHtml -notmatch 'data-benchmark-reset' -or
    $pageHtml -notmatch 'data-benchmark-sort-mode' -or
    $pageHtml -notmatch '/js/benchmarks.js') {
    throw "Benchmark page did not render matrix sorting and filtering controls."
}

$scriptText = Get-Content -LiteralPath $scriptPath -Raw -Encoding UTF8
if ($scriptText -notmatch 'OfficeImoBenchmarkMatrix' -or $scriptText -notmatch 'sortBy' -or $scriptText -notmatch 'setFilter' -or $scriptText -notmatch 'setSortMetric' -or $scriptText -notmatch 'data-ratio-to-fastest') {
    throw "Benchmark sort/filter script does not expose the expected matrix behaviors."
}

if ($scriptText -notmatch 'benchmark-workload' -or
    $scriptText -notmatch 'benchmark-os' -or
    $scriptText -notmatch 'benchmark-mode' -or
    $scriptText -notmatch "queryValue\('benchmark-os'\)" -or
    $scriptText -notmatch "queryValue\('benchmark-mode'\)" -or
    $scriptText -notmatch "filterValue\('platform'\)" -or
    $scriptText -notmatch "filterValue\('runMode'\)" -or
    $scriptText -notmatch "row\.getAttribute\('data-platform'\)" -or
    $scriptText -notmatch "row\.getAttribute\('data-run-mode'\)" -or
    $scriptText -notmatch 'candidate\.comparisonId === selectedComparison' -or
    $scriptText -notmatch 'item\.comparisonId === selectedComparison' -or
    $scriptText -notmatch 'item\.runMode === selectedMode' -or
    $scriptText -notmatch 'candidate\.publish === true' -or
    $scriptText -notmatch 'activeRequestId' -or
    $scriptText -notmatch 'requestId !== activeRequestId' -or
    $scriptText -notmatch "compatibilityValue\(entry, 'gitSha'\)" -or
    $scriptText -notmatch "sourceCommit\.substring\(0, 12\)" -or
    $scriptText -notmatch 'compatibilityIssues' -or
    $scriptText -notmatch "macos:\s*'macOS'" -or
    $scriptText -notmatch 'workloadName\(\)' -or
    $scriptText -notmatch 'comparisonGroupName\(row\)' -or
    $scriptText -notmatch "\['namespace', 'type', 'fullname'\]" -or
    $scriptText -notmatch "split\('&'\)" -or
    $scriptText -notmatch 'csv-25k-datareader-write-net10\.0' -or
    $scriptText -notmatch 'xlsx-25k-datareader-write-net10\.0' -or
    $scriptText -match "scenario === 'OfficeIMO'" -or
    $pageHtml -notmatch 'Quick results are diagnostic only') {
    throw 'Library comparison selector does not preserve shareable state, reject stale responses, and enforce evidence safety labels.'
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

$highlightCards = [regex]::Matches(
    $pageHtml,
    '<article class="imo-benchmark-highlight">.*?<tbody>(?<rows>.*?)</tbody>.*?</article>',
    [System.Text.RegularExpressions.RegexOptions]::Singleline)
foreach ($card in $highlightCards) {
    $timings = @(
        [regex]::Matches(
            $card.Groups['rows'].Value,
            '<th scope="row">(?<library>.*?)</th><td>(?<value>[0-9]+(?:\.[0-9]+)?) (?<unit>ms|s)</td>') |
            ForEach-Object {
                $milliseconds = [double]::Parse(
                    $_.Groups['value'].Value,
                    [System.Globalization.CultureInfo]::InvariantCulture)
                if ($_.Groups['unit'].Value -eq 's') {
                    $milliseconds *= 1000
                }
                [pscustomobject]@{
                    Library = [System.Net.WebUtility]::HtmlDecode($_.Groups['library'].Value)
                    Milliseconds = $milliseconds
                }
            }
    )
    for ($index = 1; $index -lt $timings.Count; $index++) {
        if ($timings[$index].Milliseconds -lt $timings[$index - 1].Milliseconds) {
            throw "Benchmark highlight rows are not fastest-first: '$($timings[$index].Library)' ($($timings[$index].Milliseconds) ms) follows '$($timings[$index - 1].Library)' ($($timings[$index - 1].Milliseconds) ms)."
        }
    }
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
