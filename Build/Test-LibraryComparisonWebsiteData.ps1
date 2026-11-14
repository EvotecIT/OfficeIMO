#requires -Version 7.4
<#
.SYNOPSIS
Checks the build-time catalog copy against existing published benchmark evidence.
.DESCRIPTION
Runs only the website generators. Uses temporary outputs to verify byte identity,
replacement of a v3 placeholder, preservation of quick entries, a write-free
DryRun, and locale-independent archive snapshots. Does not collect benchmarks
or modify published evidence.
.PARAMETER CatalogPath
Published v4 catalog containing quick evidence for the preservation check.
.EXAMPLE
./Build/Test-LibraryComparisonWebsiteData.ps1
.OUTPUTS
A confirmation of the preserved catalog and quick-entry counts.
#>
param(
    [string] $CatalogPath = (Join-Path $PSScriptRoot '../Website/static/data/benchmarks/library-comparisons/index.json')
)

$ErrorActionPreference = 'Stop'
$CatalogPath = (Resolve-Path -LiteralPath $CatalogPath).ProviderPath
$catalog = Get-Content -LiteralPath $CatalogPath -Raw -Encoding UTF8 | ConvertFrom-Json -Depth 100
$quickEntries = @($catalog.entries | Where-Object { $_.runMode -eq 'quick' -and -not $_.publish })
if ($quickEntries.Count -eq 0) {
    throw 'The input fixture must include unpublished quick results to verify their preservation.'
}
$sourceHash = (Get-FileHash -LiteralPath $CatalogPath -Algorithm SHA256).Hash
$temporaryRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('officeimo-benchmark-website-' + [guid]::NewGuid().ToString('N'))
$websiteIndex = Join-Path $temporaryRoot 'data/benchmarks_index.json'
$generator = Join-Path $PSScriptRoot 'Generate-LibraryComparisonWebsiteData.ps1'
$overview = Join-Path $PSScriptRoot 'Generate-BenchmarkWebsiteOverview.ps1'
$excelPath = Join-Path $PSScriptRoot '../Docs/benchmarks/readme-current/officeimo.excel.comparison.json'
$csvPath = Join-Path $PSScriptRoot '../Docs/benchmarks/readme-current/officeimo.csv.comparison.json'
$previousCulture = [System.Globalization.CultureInfo]::CurrentCulture
try {
    $plan = & $generator -CatalogPath $CatalogPath -WebsiteIndexPath $websiteIndex -DryRun
    if (Test-Path -LiteralPath $temporaryRoot) {
        throw 'Generator DryRun created output.'
    }
    if (@($plan.OutputPath).Count -ne 1 -or $plan.OutputPath[0] -ne $websiteIndex) {
        throw 'Generator DryRun does not expose its single build-time data destination.'
    }

    New-Item -ItemType Directory -Path (Split-Path -Parent $websiteIndex) -Force | Out-Null
    Set-Content -LiteralPath $websiteIndex -Value '{"schemaVersion":3,"entries":[]}' -Encoding utf8NoBOM
    $fallback = Get-Content -LiteralPath $excelPath -Raw -Encoding UTF8 | ConvertFrom-Json -Depth 100
    foreach ($row in $fallback.comparison) { $row.variables.PSObject.Properties.Remove('Snapshot') }
    $fallback.metadata.generatedAtUtc = '2026-08-16T10:51:57Z'
    $fallbackPath = Join-Path $temporaryRoot 'excel-date-fallback.json'
    $fallback | ConvertTo-Json -Depth 100 | Set-Content -LiteralPath $fallbackPath -Encoding utf8NoBOM

    $archiveHash = $null
    $fallbackHash = $null
    foreach ($locale in @('en-US', 'pl-PL')) {
        [System.Globalization.CultureInfo]::CurrentCulture = [System.Globalization.CultureInfo]::GetCultureInfo($locale)
        $archivePath = Join-Path $temporaryRoot "$locale/archive.html"
        $fallbackOutput = Join-Path $temporaryRoot "$locale/date-fallback.html"
        & $overview -ExcelHighlightsPath $excelPath -CsvHighlightsPath $csvPath -OutputPath $archivePath
        & $overview -ExcelHighlightsPath $fallbackPath -CsvHighlightsPath $csvPath -OutputPath $fallbackOutput
        $currentArchiveHash = (Get-FileHash -LiteralPath $archivePath).Hash
        $currentFallbackHash = (Get-FileHash -LiteralPath $fallbackOutput).Hash
        if ($archiveHash -and ($archiveHash -ne $currentArchiveHash -or $fallbackHash -ne $currentFallbackHash)) {
            throw 'The archive snapshot depends on the publishing locale.'
        }
        if ((Get-Content -LiteralPath $fallbackOutput -Raw) -notmatch 'Snapshot 2026-08-16') {
            throw 'The archive fallback did not preserve the UTC snapshot date.'
        }
        $archiveHash = $currentArchiveHash
        $fallbackHash = $currentFallbackHash

        & $generator -CatalogPath $CatalogPath -WebsiteIndexPath $websiteIndex
        if ((Get-FileHash -LiteralPath $websiteIndex).Hash -ne $sourceHash -or
            (Get-FileHash -LiteralPath $CatalogPath).Hash -ne $sourceHash -or
            (Get-FileHash -LiteralPath $archivePath).Hash -ne $archiveHash) {
            throw 'Catalog copying changed the published bytes or the separate archive snapshot.'
        }
        $mirror = Get-Content -LiteralPath $websiteIndex -Raw -Encoding UTF8 | ConvertFrom-Json -Depth 100
        $retainedQuick = @($mirror.entries | Where-Object { $_.runMode -eq 'quick' -and -not $_.publish })
        if ($retainedQuick.Count -ne $quickEntries.Count -or $mirror.entries.Count -ne $catalog.entries.Count) {
            throw 'The template index filtered benchmark evidence.'
        }
    }
    Write-Host "Website catalog and archive verified in en-US/pl-PL: $($mirror.entries.Count) entries, including $($retainedQuick.Count) unpublished quick results."
} finally {
    [System.Globalization.CultureInfo]::CurrentCulture = $previousCulture
    if (Test-Path -LiteralPath $temporaryRoot) {
        Remove-Item -LiteralPath $temporaryRoot -Recurse -ErrorAction Stop
    }
}
