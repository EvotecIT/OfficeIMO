#requires -Version 7.4
<#
.SYNOPSIS
Checks the build-time catalog copy against existing published benchmark evidence.
.DESCRIPTION
Runs only the website generator. Uses temporary outputs to verify byte identity,
replacement of a v3 placeholder, preservation of quick entries, and a write-free
DryRun. Does not collect benchmarks or modify published evidence.
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
$partialDirectory = Join-Path $temporaryRoot 'partials'
$websiteIndex = Join-Path $temporaryRoot 'data/benchmarks_index.json'
$generator = Join-Path $PSScriptRoot 'Generate-LibraryComparisonWebsiteData.ps1'
try {
    $plan = & $generator -CatalogPath $CatalogPath -OutputDirectory $partialDirectory -WebsiteIndexPath $websiteIndex -DryRun
    if (Test-Path -LiteralPath $temporaryRoot) {
        throw 'Generator DryRun created output.'
    }
    if ($websiteIndex -notin $plan.OutputPath) {
        throw 'Generator DryRun does not expose its build-time index destination.'
    }

    New-Item -ItemType Directory -Path (Split-Path -Parent $websiteIndex) -Force | Out-Null
    Set-Content -LiteralPath $websiteIndex -Value '{"schemaVersion":3,"entries":[]}' -Encoding utf8NoBOM
    & $generator -CatalogPath $CatalogPath -OutputDirectory $partialDirectory -WebsiteIndexPath $websiteIndex
    if ((Get-FileHash -LiteralPath $websiteIndex -Algorithm SHA256).Hash -ne $sourceHash -or
        (Get-FileHash -LiteralPath $CatalogPath -Algorithm SHA256).Hash -ne $sourceHash) {
        throw 'The template index differs from the published bytes, or the source was modified.'
    }
    $mirror = Get-Content -LiteralPath $websiteIndex -Raw -Encoding UTF8 | ConvertFrom-Json -Depth 100
    $retainedQuick = @($mirror.entries | Where-Object { $_.runMode -eq 'quick' -and -not $_.publish })
    if ($retainedQuick.Count -ne $quickEntries.Count -or $mirror.entries.Count -ne $catalog.entries.Count) {
        throw 'The template index filtered benchmark evidence.'
    }
    Write-Host "Website catalog copy verified: $($mirror.entries.Count) entries, including $($retainedQuick.Count) unpublished quick results."
} finally {
    if (Test-Path -LiteralPath $temporaryRoot) {
        Remove-Item -LiteralPath $temporaryRoot -Recurse -ErrorAction Stop
    }
}
