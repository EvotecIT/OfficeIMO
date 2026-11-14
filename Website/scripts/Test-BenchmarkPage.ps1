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
$archivePath = Join-Path $resolvedSiteRoot 'benchmarks\archive\index.html'
$areaCatalogPath = Join-Path $PSScriptRoot '..\data\benchmark_areas.json'
$buildIndexPath = Join-Path $PSScriptRoot '..\data\benchmarks_index.json'
$scriptPath = Join-Path $resolvedSiteRoot 'js\benchmarks.js'
$comparisonCatalogPath = Join-Path $resolvedSiteRoot 'data\benchmarks\library-comparisons\index.json'
$evidenceHelperPath = Join-Path $PSScriptRoot '..\..\Build\BenchmarkEvidence.ps1'

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

if (-not (Test-Path -LiteralPath $archivePath -PathType Leaf)) {
    throw "Benchmark archive page was not generated at '$archivePath'."
}

if (-not (Test-Path -LiteralPath $comparisonCatalogPath -PathType Leaf)) {
    throw "Library comparison evidence catalog was not published to '$comparisonCatalogPath'."
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

if ($rowCount -lt 1) {
    throw "Benchmark data JSON does not contain measurement rows."
}

if ($matrixRowCount -lt 1) {
    throw "Benchmark data JSON does not contain matrix rows."
}

if ($summaryCount -lt 1) {
    throw "Benchmark data JSON does not contain summary rows."
}

if ($comparisonCatalog.schemaVersion -ne 4 -or
    @($comparisonCatalog.expectedPlatforms).Count -ne 3 -or
    -not (@($comparisonCatalog.availability).Where({ $_.platform -eq 'macos' -and -not $_.available }))) {
    throw 'Library comparison catalog does not expose the v4 three-platform model and explicit missing macOS lane.'
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

# Every catalogued result is reachable: the overview lists its area and the area page renders a run for it.
$areaCatalog = Get-Content -LiteralPath $areaCatalogPath -Raw -Encoding UTF8 | ConvertFrom-Json
$areaIds = @($areaCatalog.areas | ForEach-Object id)
$buildIndex = Get-Content -LiteralPath $buildIndexPath -Raw -Encoding UTF8 | ConvertFrom-Json
$publishedShas = @($comparisonCatalog.entries | ForEach-Object resultSha256 | Sort-Object)
$buildShas = @($buildIndex.entries | ForEach-Object resultSha256 | Sort-Object)
if (($publishedShas -join ',') -ne ($buildShas -join ',')) {
    throw "The build-time benchmark index '$buildIndexPath' does not list the same results as the published index."
}

$overviewHtml = Get-Content -LiteralPath $pagePath -Raw -Encoding UTF8
foreach ($areaId in $areaIds) {
    if ($overviewHtml -notmatch ([regex]::Escape("data-area=""$areaId""")) ) {
        throw "The benchmark overview does not list the '$areaId' area."
    }
}

foreach ($areaGroup in @($buildIndex.entries | Group-Object area)) {
    if ([string]::IsNullOrWhiteSpace($areaGroup.Name) -or $areaGroup.Name -notin $areaIds) {
        throw "Benchmark results use area '$($areaGroup.Name)', which data/benchmark_areas.json does not define."
    }
    $areaPagePath = Join-Path $resolvedSiteRoot "benchmarks/$($areaGroup.Name)/index.html"
    if (-not (Test-Path -LiteralPath $areaPagePath -PathType Leaf)) {
        throw "Benchmark area page '/benchmarks/$($areaGroup.Name)/' was not generated."
    }
    $areaHtml = Get-Content -LiteralPath $areaPagePath -Raw -Encoding UTF8
    foreach ($entry in $areaGroup.Group) {
        if ($areaHtml -notmatch ([regex]::Escape("id=""$($entry.comparisonId)""")) -or
            $areaHtml -notmatch ([regex]::Escape("data-result=""$($entry.resultPath)"""))) {
            throw "Benchmark area page '/benchmarks/$($areaGroup.Name)/' does not render result '$($entry.resultPath)'."
        }
    }
}

# The archive keeps the complete historical Excel matrix.
$archiveHtml = Get-Content -LiteralPath $archivePath -Raw -Encoding UTF8
$renderedMatrixRows = ([regex]::Matches($archiveHtml, '<tr[^>]*data-benchmark-row[^>]*>\s*<td class="imo-benchmark-scenario"[^>]*data-label="Scenario"')).Count
if ($renderedMatrixRows -lt $matrixRowCount) {
    throw "Benchmark archive rendered $renderedMatrixRows matrix rows, expected at least $matrixRowCount."
}
$sortableCells = ([regex]::Matches($archiveHtml, 'data-library="[^"]+"')).Count
if ($sortableCells -lt $rowCount) {
    throw "Benchmark archive rendered $sortableCells sortable library cells, expected at least $rowCount."
}

& (Join-Path $PSScriptRoot 'Test-BenchmarkPage.Browser.ps1') -SiteRoot $resolvedSiteRoot

Write-Host "Benchmark pages verified: $(@($buildIndex.entries).Count) results across $(@($buildIndex.entries | Group-Object area).Count) areas, $matrixRowCount archived matrix rows."
