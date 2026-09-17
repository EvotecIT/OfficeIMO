$ErrorActionPreference = 'Stop'

$runner = Join-Path $PSScriptRoot 'Run-LibraryComparisonBenchmarks.ps1'
$comparisonProject = Join-Path (Split-Path -Parent $PSScriptRoot) `
    'OfficeIMO.Pdf.Benchmarks.Comparisons\OfficeIMO.Pdf.Benchmarks.Comparisons.csproj'
[xml] $comparisonProjectXml = Get-Content -LiteralPath $comparisonProject -Raw
$questPdfReference = @($comparisonProjectXml.Project.ItemGroup.PackageReference) |
    Where-Object Include -eq 'QuestPDF' |
    Select-Object -First 1
$questPdfVersionProperty = @($comparisonProjectXml.SelectNodes('/Project/PropertyGroup/QuestPdfBenchmarkVersion')) |
    Where-Object { $_.GetAttribute('Condition') -eq "'`$(QuestPdfBenchmarkVersion)' == ''" } |
    Select-Object -First 1
$questPdfBoundaryTarget = $comparisonProjectXml.SelectSingleNode(
    "/Project/Target[@Name='ValidateQuestPdfBenchmarkLicenseBoundary']")
if ($null -eq $questPdfReference -or [string] $questPdfReference.Version -ne '$(QuestPdfBenchmarkVersion)' -or
    $null -eq $questPdfVersionProperty -or $questPdfVersionProperty.InnerText -ne '2026.5.0' -or
    $null -eq $questPdfBoundaryTarget -or
    [string] $questPdfBoundaryTarget.Error.Condition -notmatch 'QuestPdfInternalAuthorization') {
    throw 'QuestPDF comparison workloads are not pinned to the last MIT-compatible package release.'
}

$publicQuestPdf = @(
    & $runner -Workload pdfgenerate -RunMode quick -PlanOnly
)
if ($publicQuestPdf.Count -ne 1 -or
    $publicQuestPdf[0].QuestPdfPackageVersion -ne '2026.5.0' -or
    $publicQuestPdf[0].InternalOnly -or
    -not $publicQuestPdf[0].CatalogEligible) {
    throw 'The public QuestPDF lane is not pinned, labelled, and catalog eligible under the MIT-compatible release.'
}

$domainPlan = @(
    & $runner -Workload pdfgenerate -RunMode full -AffinityMask 4294901760 -PlanOnly
)
if ($domainPlan.Count -ne 1 -or
    $domainPlan[0].CatalogComparisonId -ne 'pdf-structured-generation-net10.0-affinity-0xffff0000' -or
    $domainPlan[0].ProvenanceWorkloadId -ne $domainPlan[0].CatalogComparisonId -or
    $domainPlan[0].AffinityMask -ne '0xFFFF0000') {
    throw 'CPU-domain evidence does not receive one stable, non-colliding catalog and provenance identity.'
}

$internalQuestPdf = @(
    & $runner `
        -Workload pdfgenerate `
        -RunMode full `
        -QuestPdfPackageVersion 2026.9.0 `
        -QuestPdfLicenseType Evaluation `
        -InternalQuestPdf `
        -ConfirmQuestPdfAuthorization `
        -PlanOnly
)
if ($internalQuestPdf.Count -ne 1 -or
    $internalQuestPdf[0].QuestPdfPackageVersion -ne '2026.9.0' -or
    -not $internalQuestPdf[0].InternalOnly -or
    $internalQuestPdf[0].CatalogEligible -or
    $internalQuestPdf[0].WillCatalog -or
    $internalQuestPdf[0].Publish) {
    throw 'The internal QuestPDF lane can still reach shared benchmark publication.'
}

$internalPublishBlocked = $false
try {
    & $runner `
        -Workload pdfgenerate `
        -RunMode full `
        -QuestPdfPackageVersion 2026.9.0 `
        -InternalQuestPdf `
        -ConfirmQuestPdfAuthorization `
        -Publish `
        -PlanOnly | Out-Null
} catch {
    $internalPublishBlocked = $true
}
if (-not $internalPublishBlocked) {
    throw 'Internal QuestPDF mode can be combined with benchmark publication.'
}

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

$affinityBeforePlan = if ([System.Runtime.InteropServices.RuntimeInformation]::IsOSPlatform(
        [System.Runtime.InteropServices.OSPlatform]::Windows)) {
    [System.Diagnostics.Process]::GetCurrentProcess().ProcessorAffinity.ToInt64()
} else {
    $null
}
$affinityPlan = @(
    & $runner -Workload csv -RunMode quick -AffinityMask 3 -PlanOnly
)
$affinityAfterPlan = if ([System.Runtime.InteropServices.RuntimeInformation]::IsOSPlatform(
        [System.Runtime.InteropServices.OSPlatform]::Windows)) {
    [System.Diagnostics.Process]::GetCurrentProcess().ProcessorAffinity.ToInt64()
} else {
    $null
}
$expectedAffinityApplication = if ([System.Runtime.InteropServices.RuntimeInformation]::IsOSPlatform(
        [System.Runtime.InteropServices.OSPlatform]::Windows)) {
    'inherited-parent-process'
} else {
    'benchmarkdotnet-command-line'
}
if ($affinityPlan.Count -ne 1 -or
    $affinityPlan[0].AffinityMask -ne '0x3' -or
    $affinityPlan[0].AffinityApplication -ne $expectedAffinityApplication -or
    $affinityAfterPlan -ne $affinityBeforePlan) {
    throw 'The benchmark plan does not report the platform-specific affinity application mechanism.'
}
if ($null -ne $affinityBeforePlan) {
    $availableAffinity = [UInt64] $affinityBeforePlan
    $singleProcessorMask = [UInt64] 1
    while (($availableAffinity -band $singleProcessorMask) -eq 0) {
        $singleProcessorMask = $singleProcessorMask -shl 1
    }
    $failedAsExpected = $false
    try {
        & $runner `
            -Workload csv `
            -RunMode quick `
            -AffinityMask $singleProcessorMask `
            -PowerForgeRoot (Join-Path ([System.IO.Path]::GetTempPath()) ([Guid]::NewGuid().ToString('N'))) |
            Out-Null
    } catch {
        $failedAsExpected = $true
    }
    $affinityAfterFailure = [System.Diagnostics.Process]::GetCurrentProcess().ProcessorAffinity.ToInt64()
    if (-not $failedAsExpected -or $affinityAfterFailure -ne $affinityBeforePlan) {
        throw 'The benchmark runner does not restore the invoking process affinity after a failed real run.'
    }
}

$pdfStructuredRead = @(
    & $runner -Workload pdfstructuredread -RunMode quick -PlanOnly
)
if ($pdfStructuredRead.Count -ne 1 -or
    $pdfStructuredRead[0].Filter -ne '*PdfStructuredRead*Benchmarks*' -or
    $pdfStructuredRead[0].ComparisonId -ne 'officeimo-pdf-read-profile-route-health-net10.0' -or
    $pdfStructuredRead[0].ExpectedCaseCount -ne 6 -or
    $pdfStructuredRead[0].CatalogEligible) {
    throw 'The canonical PDF read profiles are not isolated as non-catalog route-health measurements.'
}

Write-Host 'Library comparison runner policy verified for standalone diagnostics, canonical PDF structured read, HTML-to-PDF route health, and mixed comparison selection.'
