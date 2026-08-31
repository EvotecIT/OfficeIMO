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
$epubRead = @($mixed | Where-Object Workload -eq 'epubread')
$zipTraversal = @($mixed | Where-Object Workload -eq 'ziptraverse')
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
if ($epubRead.Count -ne 1 -or
    -not $epubRead[0].CatalogEligible -or -not $epubRead[0].WillCatalog -or
    -not $epubRead[0].Publish) {
    throw 'The EPUB open/read comparison is missing from the catalog-eligible mixed benchmark plan.'
}
if ($zipTraversal.Count -ne 1 -or
    $zipTraversal[0].CatalogEligible -or
    $zipTraversal[0].WillCatalog -or
    $zipTraversal[0].Publish) {
    throw 'The non-equivalent ZIP safety-overhead diagnostic can reach library-comparison publication.'
}

$pdfHtml = @(
    & $runner -Workload pdfhtml -RunMode quick -PlanOnly
)
if ($pdfHtml.Count -ne 1 -or
    $pdfHtml[0].ComparisonId -ne 'pdf-html-generation-net10.0' -or
    $pdfHtml[0].ExpectedCaseCount -ne 12) {
    throw 'The HTML-to-PDF comparison plan does not require all four engines at all three scales.'
}

$unrelated = @(
    & $runner -Workload csv -RunMode quick -HtmlTinkerXRoot 'missing-htmltinkerx-that-must-not-be-resolved' -PlanOnly
)
if ($unrelated.Count -ne 1 -or $unrelated[0].Workload -ne 'csv') {
    throw 'An unrelated comparison workload still depends on HtmlTinkerX discovery.'
}

$pdfStructuredRead = @(
    & $runner -Workload pdfstructuredread -RunMode quick -PlanOnly
)
if ($pdfStructuredRead.Count -ne 1 -or
    $pdfStructuredRead[0].Filter -ne '*PdfStructuredReadBenchmarks*' -or
    $pdfStructuredRead[0].ComparisonId -ne 'officeimo-pdf-read-profile-route-health-net10.0' -or
    $pdfStructuredRead[0].ExpectedCaseCount -ne 6 -or
    $pdfStructuredRead[0].CatalogEligible) {
    throw 'The canonical PDF read profiles are not isolated as non-catalog route-health measurements.'
}

Write-Host 'Library comparison runner policy verified for standalone diagnostics, canonical PDF structured read, HTML-to-PDF route health, and mixed comparison selection.'
