$ErrorActionPreference = 'Stop'

$runner = Join-Path $PSScriptRoot 'Run-LibraryComparisonBenchmarks.ps1'

$standalone = @(
    & $runner -Workload pdfformats -RunMode full -Publish -PlanOnly
)
if ($standalone.Count -ne 1 -or
    $standalone[0].Workload -ne 'pdfformats' -or
    $standalone[0].ComparisonId -ne 'officeimo-pdf-format-route-health-net10.0' -or
    $standalone[0].CatalogEligible -or
    $standalone[0].WillCatalog -or
    $standalone[0].Publish) {
    throw 'The standalone PDF format route-health workload can reach library-comparison publication.'
}

$mixed = @(
    & $runner -Workload all -RunMode full -Publish -PlanOnly
)
$mixedRouteHealth = @($mixed | Where-Object Workload -eq 'pdfformats')
$mixedComparison = @($mixed | Where-Object Workload -eq 'pdfhtmlpayload')
$odsCreate = @($mixed | Where-Object Workload -eq 'odscreate')
$odsRead = @($mixed | Where-Object Workload -eq 'odsread')
if ($mixedRouteHealth.Count -ne 1 -or
    $mixedRouteHealth[0].CatalogEligible -or
    $mixedRouteHealth[0].WillCatalog -or
    $mixedRouteHealth[0].Publish) {
    throw 'The mixed benchmark plan can publish PDF format route-health evidence as a library comparison.'
}
if ($mixedComparison.Count -ne 1 -or
    -not $mixedComparison[0].CatalogEligible -or
    -not $mixedComparison[0].WillCatalog -or
    -not $mixedComparison[0].Publish) {
    throw 'The mixed benchmark plan no longer publishes eligible library-comparison evidence.'
}
if ($odsCreate.Count -ne 1 -or $odsRead.Count -ne 1 -or
    -not $odsCreate[0].CatalogEligible -or -not $odsCreate[0].WillCatalog -or
    -not $odsRead[0].CatalogEligible -or -not $odsRead[0].WillCatalog) {
    throw 'The ODS create/read comparisons are missing from the catalog-eligible mixed benchmark plan.'
}

Write-Host 'Library comparison runner policy verified for standalone and mixed PDF route-health selection.'
