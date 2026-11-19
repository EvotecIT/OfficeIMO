$ErrorActionPreference = 'Stop'

$runner = Join-Path $PSScriptRoot 'Run-LibraryComparisonBenchmarks.ps1'
$comparisonProject = Join-Path (Split-Path -Parent $PSScriptRoot) `
    'OfficeIMO.Pdf.Benchmarks.Comparisons\OfficeIMO.Pdf.Benchmarks.Comparisons.csproj'
$pdfGenerationBenchmark = Join-Path (Split-Path -Parent $PSScriptRoot) `
    'OfficeIMO.Pdf.Benchmarks.Comparisons\PdfGenerationBenchmarks.cs'
[xml] $comparisonProjectXml = Get-Content -LiteralPath $comparisonProject -Raw
$pdfGenerationBenchmarkText = Get-Content -LiteralPath $pdfGenerationBenchmark -Raw
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
if ($pdfGenerationBenchmarkText -notmatch
    '(?s)\[GlobalSetup\]\s*public void Setup\(\)\s*\{\s*BenchmarkAffinityGuard\.Validate\(\);') {
    throw 'The published structured-PDF benchmark does not validate affinity inside its measured worker.'
}

$publicQuestPdf = @(
    & $runner -Workload pdfgenerate -Profile diagnostic -PlanOnly
)
if ($publicQuestPdf.Count -ne 1 -or
    $publicQuestPdf[0].QuestPdfPackageVersion -ne '2026.5.0' -or
    $publicQuestPdf[0].InternalOnly -or
    -not $publicQuestPdf[0].CatalogEligible -or
    $publicQuestPdf[0].ComparisonId -ne 'pdf-structured-generation') {
    throw 'The public QuestPDF lane is not pinned, labelled, and catalog eligible under the MIT-compatible release.'
}

$domainPlan = @(
    & $runner -Workload pdfgenerate -Profile full -AffinityMask 4294901760 -PlanOnly
)
# Processor affinity is a lane parameter of the same comparison, never part of its id.
if ($domainPlan.Count -ne 1 -or
    $domainPlan[0].ComparisonId -ne 'pdf-structured-generation' -or
    $domainPlan[0].ProvenanceWorkloadId -ne $domainPlan[0].ComparisonId -or
    $domainPlan[0].LaneParameters.affinity -ne '0xffff0000' -or
    $domainPlan[0].AffinityMask -ne '0xFFFF0000') {
    throw 'CPU-domain evidence must be a separate lane of the same comparison, with affinity as a lane parameter.'
}

$internalQuestPdf = @(
    & $runner `
        -Workload pdfgenerate `
        -Profile full `
        -QuestPdfPackageVersion 2026.9.0 `
        -AcceptLibraryLicense questpdf-latest `
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

$unacceptedQuestPdfBlocked = $false
try {
    & $runner -Workload pdfgenerate -Profile diagnostic -QuestPdfPackageVersion 2026.9.0 -PlanOnly | Out-Null
} catch {
    $unacceptedQuestPdfBlocked = $true
}
if (-not $unacceptedQuestPdfBlocked) {
    throw 'A newer QuestPDF release can run without accepting its licence.'
}

$internalPublishBlocked = $false
try {
    & $runner `
        -Workload pdfgenerate `
        -Profile full `
        -QuestPdfPackageVersion 2026.9.0 `
        -AcceptLibraryLicense questpdf-latest `
        -Publish `
        -PlanOnly | Out-Null
} catch {
    $internalPublishBlocked = $true
}
if (-not $internalPublishBlocked) {
    throw 'Internal QuestPDF mode can be combined with benchmark publication.'
}

$standalone = @(
    & $runner -Workload pdfformats -Profile full -Publish -PlanOnly
)
if ($standalone.Count -ne 1 -or
    $standalone[0].Workload -ne 'pdfformats' -or
    $standalone[0].ComparisonId -ne 'officeimo-pdf-format-route-health' -or
    $standalone[0].CatalogEligible -or
    $standalone[0].WillCatalog -or
    $standalone[0].Publish) {
    throw 'The standalone PDF format route-health workload can reach library-comparison publication.'
}

$mixed = @(
    & $runner -Workload all -Profile full -Publish -PlanOnly
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
    -not $zipTraversal[0].WillCatalog -or
    $zipTraversal[0].Publish -or
    [string]::IsNullOrWhiteSpace($zipTraversal[0].PublishRestriction)) {
    throw 'ZIP safety-overhead evidence must remain visible as a diagnostic without publication.'
}

$pdfHtml = @(
    & $runner -Workload pdfhtml -Profile diagnostic -PlanOnly
)
if ($pdfHtml.Count -ne 1 -or
    $pdfHtml[0].ComparisonId -ne 'pdf-html-generation' -or
    $pdfHtml[0].ExpectedCaseCount -ne 15) {
    throw 'The HTML-to-PDF comparison plan does not require all five engines at all three scales.'
}

$pdfHtmlPayload = @(& $runner -Workload pdfhtmlpayload -Profile diagnostic -PlanOnly)
if ($pdfHtmlPayload.Count -ne 1 -or $pdfHtmlPayload[0].ExpectedCaseCount -ne 9) {
    throw 'The exact 21 KiB HTML comparison plan does not require all three engines and payloads.'
}

$unrelated = @(
    & $runner -Workload csv -Profile diagnostic -HtmlTinkerXRoot 'missing-htmltinkerx-that-must-not-be-resolved' -PlanOnly
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
    & $runner -Workload csv -Profile diagnostic -AffinityMask 3 -PlanOnly
)
$affinityAfterPlan = if ([System.Runtime.InteropServices.RuntimeInformation]::IsOSPlatform(
        [System.Runtime.InteropServices.OSPlatform]::Windows)) {
    [System.Diagnostics.Process]::GetCurrentProcess().ProcessorAffinity.ToInt64()
} else {
    $null
}
$inheritsAffinity = [System.Runtime.InteropServices.RuntimeInformation]::IsOSPlatform(
    [System.Runtime.InteropServices.OSPlatform]::Windows) -or
    [System.Runtime.InteropServices.RuntimeInformation]::IsOSPlatform(
        [System.Runtime.InteropServices.OSPlatform]::Linux)
$expectedAffinityApplication = if ($inheritsAffinity) {
    'inherited-parent-process'
} else {
    'unsupported'
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
            -Profile diagnostic `
            -AffinityMask $singleProcessorMask `
            -PowerForgeRoot (Split-Path -Parent $PSScriptRoot) |
            Out-Null
    } catch {
        if ($_.Exception.Message -notlike 'PowerForgeRoot must be at the pinned commit *') {
            throw
        }
        $failedAsExpected = $true
    }
    $affinityAfterFailure = [System.Diagnostics.Process]::GetCurrentProcess().ProcessorAffinity.ToInt64()
    if (-not $failedAsExpected -or $affinityAfterFailure -ne $affinityBeforePlan) {
        throw 'The benchmark runner does not restore the invoking process affinity after rejected setup.'
    }
}

$pdfStructuredRead = @(
    & $runner -Workload pdfstructuredread -Profile diagnostic -PlanOnly
)
if ($pdfStructuredRead.Count -ne 1 -or
    $pdfStructuredRead[0].Filter -ne '*PdfStructuredRead*Benchmarks*' -or
    $pdfStructuredRead[0].ComparisonId -ne 'officeimo-pdf-read-profile-route-health' -or
    $pdfStructuredRead[0].ExpectedCaseCount -ne 6 -or
    $pdfStructuredRead[0].CatalogEligible) {
    throw 'The canonical PDF read profiles are not isolated as non-catalog route-health measurements.'
}

# The declarations the catalog checks results against are generated from the runner definitions and the library
# policy, and are current.
& (Join-Path $PSScriptRoot 'Export-BenchmarkDeclarations.ps1') -Check
$declarations = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'benchmark-declarations.json') -Raw | ConvertFrom-Json
foreach ($comparison in @($declarations.comparisons)) {
    if ($comparison.comparisonId -match '-net\d|-affinity-') {
        throw "Comparison id '$($comparison.comparisonId)' encodes a lane dimension."
    }
}
# A single-library comparison is recorded but never published, and a framework on which a library runs another
# version is not declared, so its lanes cannot run.
$singleLibrary = @(& $runner -Workload pdfnative -Profile full -Publish -PlanOnly)
if ($singleLibrary[0].Publish -or -not $singleLibrary[0].WillCatalog) {
    throw 'A single-library comparison must be recorded without publication.'
}
$net8Xlsx = @(& $runner -Workload xlsx -Profile full -Framework net8.0 -PlanOnly)
if ($net8Xlsx[0].CanRun -or 'net8.0' -in @($net8Xlsx[0].DeclaredFrameworks)) {
    throw 'XLSX reading must not run on net8.0, where ExcelReader.NET resolves another version than the declared one.'
}
$diagnosticPublishBlocked = $false
try {
    & $runner -Workload csv -Profile diagnostic -Publish -PlanOnly | Out-Null
} catch {
    $diagnosticPublishBlocked = $true
}
if (-not $diagnosticPublishBlocked) {
    throw 'A diagnostic run must never be publishable.'
}

$odsPlan = @(& $runner -Workload odscreate -Profile full -PlanOnly)
if ('OpenStandardLibrary|Scale=Normal' -notin @($odsPlan[0].ExpectedCases)) {
    throw 'The plan must list every expected method and parameter combination.'
}

# The library policy declares every compared library, and only open libraries reach a published lane.
. (Join-Path $PSScriptRoot 'LibraryComparisonDefinitions.ps1')
. (Join-Path $PSScriptRoot 'BenchmarkLibraryPolicy.ps1')
$policy = Get-BenchmarkLibraryPolicy
foreach ($library in @($policy.libraries)) {
    $hasLicense = $library.license -and ($library.license.expression -or $library.license.name) -and $library.license.url
    if (-not $hasLicense -or [string]::IsNullOrWhiteSpace($library.reason) -or $null -eq $library.packages) {
        throw "Benchmark library '$($library.id)' needs packages, a licence with a URL and a reason."
    }
    if ($library.restriction -eq 'acceptance-required' -and $null -eq $library.acceptance) {
        throw "Benchmark library '$($library.id)' requires acceptance but does not say what the caller supplies."
    }
}
foreach ($framework in @('net8.0', 'net10.0')) {
    foreach ($entry in (Get-LibraryComparisonDefinitions -Framework $framework).GetEnumerator()) {
        foreach ($name in @($entry.Value.Libraries)) {
            $null = Resolve-BenchmarkLibrary -Policy $policy -Name $name
        }
    }
}
foreach ($framework in @('net8.0', 'net10.0')) {
    $publishedPlan = @(& $runner -Workload all -Profile full -Framework $framework -Publish -PlanOnly)
    foreach ($item in @($publishedPlan | Where-Object Publish)) {
        $restricted = @($item.Libraries | ForEach-Object { Resolve-BenchmarkLibrary -Policy $policy -Name $_ } |
            Where-Object restriction -ne 'open')
        $restrictedCases = @($item.ExpectedCases | Where-Object {
                $case = $_
                @($policy.libraries | Where-Object { $_.restriction -ne 'open' -and (Test-BenchmarkCaseMatchesLibrary -Library $_ -Case $case) }).Count -gt 0
            })
        if ($restricted.Count -gt 0 -or $restrictedCases.Count -gt 0 -or $item.Environment.OFFICEIMO_BENCHMARK_ENABLED_LIBRARIES) {
            throw "Published lane '$($item.Workload)' ($framework) can include a restricted library."
        }
    }
}
# A published Word lane drops the acceptance-required DocX cases and publishes OfficeIMO, NPOI and Open XML SDK.
$wordPlan = @(& $runner -Workload wordcreate -Profile full -Framework net8.0 -Publish -PlanOnly)
if (-not $wordPlan[0].Publish -or 'DocX' -in @($wordPlan[0].Libraries) -or
    @($wordPlan[0].ExpectedCases | Where-Object { $_ -match '^DocX\|' }).Count -gt 0 -or
    @($wordPlan[0].ExpectedCases | Where-Object { $_ -match '^NPOI\|' }).Count -eq 0) {
    throw 'Word comparisons must publish without DocX cases and with NPOI.'
}
$docxIncludeRejected = $false
try {
    & $runner -Workload wordcreate -Profile full -Framework net8.0 -IncludeLibrary docx -PlanOnly | Out-Null
} catch {
    $docxIncludeRejected = $true
}
if (-not $docxIncludeRejected) {
    throw 'DocX must need -AcceptLibraryLicense docx, not only -IncludeLibrary.'
}
$localDocX = @(& $runner -Workload wordcreate -Profile full -Framework net8.0 -AcceptLibraryLicense docx -PlanOnly)
if ($localDocX[0].CatalogEligible -or $localDocX[0].Publish -or $localDocX[0].Environment.OFFICEIMO_BENCHMARK_ENABLED_LIBRARIES -ne 'docx') {
    throw 'Accepting DocX must keep the Word run local.'
}
# An explicitly enabled restricted library is required, not merely tolerated: its cases join the expected set and
# its own licence value is checked before the run.
$previousEpPlusLicense = $env:EPPlusLicense
try {
    $env:EPPlusLicense = $null
    $defaultXlsx = @(& $runner -Workload xlsx -Profile full -PlanOnly)
    $enabledXlsx = @(& $runner -Workload xlsx -Profile full -AcceptLibraryLicense epplus -PlanOnly)
} finally {
    $env:EPPlusLicense = $previousEpPlusLicense
}
if ('EPPlus' -in @($defaultXlsx[0].ExpectedCases) -or 'EPPlus' -in @($defaultXlsx[0].Libraries) -or
    'EPPlus' -notin @($enabledXlsx[0].ExpectedCases) -or
    @($enabledXlsx[0].BlockedReasons | Where-Object { $_ -match 'EPPlusLicense' }).Count -eq 0) {
    throw 'Enabling EPPlus must require its XLSX case and its own licence declaration; leaving it out must drop both.'
}

# -LocalOnly keeps a run, even a pinned QuestPDF one, out of the shared catalog; the QuestPDF wrapper always uses it.
$localPinned = @(& $runner -Workload pdfgenerate -Profile diagnostic -QuestPdfPackageVersion 2026.5.0 -LocalOnly -PlanOnly)
if ($localPinned[0].WillCatalog -or $localPinned[0].Publish -or $null -ne $localPinned[0].EvidenceCatalog -or
    -not $localPinned[0].InternalOnly) {
    throw '-LocalOnly must keep the run out of the shared catalog.'
}
$wrapperAst = [System.Management.Automation.Language.Parser]::ParseFile(
    (Join-Path $PSScriptRoot 'Run-InternalQuestPdfBenchmarks.ps1'), [ref] $null, [ref] $null)
$wrapperCall = @($wrapperAst.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.CommandAst] -and $node.InvocationOperator -eq 'Ampersand' -and
            $node.CommandElements[0].Extent.Text -match 'Run-LibraryComparisonBenchmarks\.ps1'
        }, $true))
if ($wrapperCall.Count -ne 1 -or
    @($wrapperCall[0].CommandElements | Where-Object { $_.Extent.Text -eq '-LocalOnly' }).Count -ne 1 -or
    @($wrapperCall[0].CommandElements | Where-Object { $_.Extent.Text -eq '-EvidenceRoot' }).Count -ne 1) {
    throw 'The QuestPDF wrapper must run the comparison with -LocalOnly and its own evidence root.'
}

$restrictedPublishBlocked = $false
try {
    & $runner -Workload xlsx -Profile full -Publish -AcceptLibraryLicense epplus -PlanOnly | Out-Null
} catch {
    $restrictedPublishBlocked = $true
}
if (-not $restrictedPublishBlocked) {
    throw 'An enabled restricted library can be combined with benchmark publication.'
}
$repositoryRoot = Split-Path -Parent $PSScriptRoot
foreach ($tag in @(
        @{ Path = 'OfficeIMO.Excel.Benchmarks\MarkPflug65KExcelBenchmarks.cs'; Pattern = 'CategoryPrefix \+ "epplus"' },
        @{ Path = 'OfficeIMO.Excel.Benchmarks\ExcelLibraryComparisonRunner.cs'; Pattern = '\["EPPlus"\] = "epplus"' },
        @{ Path = 'OfficeIMO.Word.Benchmarks\WordCreateComparisonBenchmarks.cs'; Pattern = 'Library:docx' },
        @{ Path = 'OfficeIMO.Word.Benchmarks\WordCreateComparisonBenchmarks.cs'; Pattern = 'IsEnabled\("docx"\)\) WordBenchmarkCorpus' })) {
    if ((Get-Content -LiteralPath (Join-Path $repositoryRoot $tag.Path) -Raw) -notmatch $tag.Pattern) {
        throw "$($tag.Path) no longer marks its restricted library cases ($($tag.Pattern))."
    }
}
# The measured process must start only after restricted-library settings are applied: the enabled list is set and
# values belonging to libraries left out of the run are cleared, in the same try block, before dotnet launches.
$runnerAst = [System.Management.Automation.Language.Parser]::ParseFile($runner, [ref] $null, [ref] $null)
$measureLaunch = @($runnerAst.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.CommandAst] -and
            $node.CommandElements.Count -ge 2 -and $node.CommandElements[0].Extent.Text -eq 'dotnet' -and
            $node.CommandElements[1].Extent.Text -eq '@arguments'
        }, $true))
if ($measureLaunch.Count -ne 1) { throw 'The runner no longer has exactly one measured dotnet launch.' }
$measureTry = $measureLaunch[0].Parent
while ($null -ne $measureTry -and $measureTry -isnot [System.Management.Automation.Language.TryStatementAst]) { $measureTry = $measureTry.Parent }
if ($null -eq $measureTry) { throw 'The measured dotnet launch is no longer inside the environment try block.' }
$launchOffset = $measureLaunch[0].Extent.StartOffset
$clears = @($measureTry.Body.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and
            $node.Member.Extent.Text -eq 'SetEnvironmentVariable' -and $node.Arguments.Count -eq 2 -and
            $node.Arguments[1].Extent.Text -eq '$null'
        }, $true))
$enables = @($measureTry.Body.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.AssignmentStatementAst] -and
            $node.Left.Extent.Text -eq '$env:OFFICEIMO_BENCHMARK_ENABLED_LIBRARIES'
        }, $true))
if ($clears.Count -eq 0 -or $enables.Count -eq 0 -or
    @($clears + $enables | Where-Object { $_.Extent.StartOffset -gt $launchOffset }).Count -gt 0) {
    throw 'The runner must set the enabled libraries and clear left-out licence values before launching dotnet.'
}
$pinnedQuestPdf = @(& $runner -Workload pdfgenerate -Profile diagnostic -PlanOnly)
if ($pinnedQuestPdf[0].Environment.Contains('OFFICEIMO_QUESTPDF_LICENSE_TYPE')) {
    throw 'A pinned QuestPDF run must keep the caller''s OFFICEIMO_QUESTPDF_LICENSE_TYPE.'
}
$xlsxPlan = @(& $runner -Workload xlsx -Profile diagnostic -PlanOnly)
if (-not $xlsxPlan[0].Environment.Contains('EPPlusLicense') -or $null -ne $xlsxPlan[0].Environment['EPPlusLicense']) {
    throw 'The plan must show that EPPlusLicense is cleared when EPPlus is not enabled.'
}

# Licence acceptance is a per-run decision, never a project default.
$acceptingProjects = @(Get-ChildItem -LiteralPath $repositoryRoot -Recurse -Filter '*.csproj' -File |
    Where-Object { $_.FullName -notmatch '[\\/](bin|obj|Ignore|artifacts)[\\/]' } |
    Where-Object { (Get-Content -LiteralPath $_.FullName -Raw) -match '<AcceptNPOIOSMFLicense>\s*true' })
if ($acceptingProjects.Count -gt 0) {
    throw "Project files accept NPOI's OSMF EULA by default: $($acceptingProjects.Name -join ', ')."
}

# Library and case categories in benchmark code must name what the declarations derive.
& (Join-Path $PSScriptRoot 'Test-BenchmarkCategories.ps1') | Out-Null

Write-Host 'Library comparison runner policy verified for the library policy, standalone diagnostics, canonical PDF structured read, HTML-to-PDF route health, and mixed comparison selection.'
