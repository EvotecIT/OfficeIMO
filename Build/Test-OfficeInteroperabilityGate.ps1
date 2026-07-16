param(
    [string] $Configuration = "Debug",
    [string] $Framework = "net8.0",
    [ValidateSet("Full", "Corpus", "Safety")]
    [string] $Suite = "Full",
    [switch] $NoRestore,
    [switch] $NoBuild
)

$ErrorActionPreference = 'Stop'
$PSNativeCommandUseErrorActionPreference = $true

$repoRoot = Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')
$projects = @{
    Drawing = Join-Path $repoRoot 'OfficeIMO.Drawing.Tests/OfficeIMO.Drawing.Tests.csproj'
    Excel = Join-Path $repoRoot 'OfficeIMO.Excel.Tests/OfficeIMO.Excel.Tests.csproj'
    Word = Join-Path $repoRoot 'OfficeIMO.Word.Tests/OfficeIMO.Word.Tests.csproj'
}

foreach ($project in $projects.GetEnumerator()) {
    if (-not (Test-Path -LiteralPath $project.Value)) {
        throw "OfficeIMO $($project.Key) test project was not found: $($project.Value)"
    }
}

function Invoke-InteroperabilityGateStep {
    param(
        [Parameter(Mandatory)]
        [string] $Name,

        [Parameter(Mandatory)]
        [string] $Project,

        [Parameter(Mandatory)]
        [string] $Filter
    )

    Write-Host ""
    Write-Host "== $Name ==" -ForegroundColor Cyan
    Write-Host "Filter: $Filter" -ForegroundColor DarkCyan

    $arguments = @(
        'test',
        $Project,
        '--configuration', $Configuration,
        '--framework', $Framework,
        '--filter', $Filter,
        '--logger', 'console;verbosity=normal'
    )
    if ($NoRestore) {
        $arguments += '--no-restore'
    }
    if ($NoBuild) {
        $arguments += '--no-build'
    }

    $started = [DateTime]::UtcNow
    Push-Location $repoRoot
    try {
        & dotnet @arguments
        if ($LASTEXITCODE -ne 0) {
            throw "$Name failed with exit code $LASTEXITCODE."
        }
    } finally {
        Pop-Location
    }

    $elapsed = [DateTime]::UtcNow - $started
    Write-Host ("Completed {0} in {1:mm\:ss}." -f $Name, $elapsed) -ForegroundColor Green
}

$excelLegacyCorpusFilter = @(
    'FullyQualifiedName=OfficeIMO.Tests.Excel.LegacyXls_Corpus_Fixtures_MatchApprovedImportReports',
    'FullyQualifiedName=OfficeIMO.Tests.Excel.LegacyXls_DiagnosticCorpus_Fixtures_MatchApprovedImportReports',
    'FullyQualifiedName=OfficeIMO.Tests.Excel.LegacyXls_Corpus_ProjectionGapSummary_MatchesApprovedBaseline'
) -join '|'

$xlsbExcelAuthoredFilter = @(
    'FullyQualifiedName~OfficeIMO.Tests.Excel.Xlsb_ExcelGeneratedFixture_',
    'FullyQualifiedName~OfficeIMO.Tests.Excel.Xlsb_StyledExcelFixture_',
    'FullyQualifiedName~OfficeIMO.Tests.Excel.Xlsb_GeometryFixture_',
    'FullyQualifiedName~OfficeIMO.Tests.Excel.Xlsb_HyperlinkFixture_',
    'FullyQualifiedName~OfficeIMO.Tests.Excel.Xlsb_RichFormulaFixture_'
) -join '|'

$excelAdvancedFilter = @(
    'FullyQualifiedName~OfficeIMO.Tests.Excel.Compatibility_Corpus_PreserveOnlyMetadata_',
    'FullyQualifiedName~OfficeIMO.Tests.Excel.Compatibility_Corpus_MacroEmbeddedObjectsAndControls_',
    'FullyQualifiedName~OfficeIMO.Tests.Excel.PackagePayloads_EmbeddedPackageCanBeHashedReplacedAndRemoved',
    'FullyQualifiedName~OfficeIMO.Tests.Excel.ThreadedComments_',
    'FullyQualifiedName~OfficeIMO.Tests.Excel.FormulaEvaluator_CalculatesRegisteredCustomFunctions',
    'FullyQualifiedName~OfficeIMO.Tests.Excel.FeatureReport_Preflight_AllowsCleanReportWorkbookWorkflows',
    'FullyQualifiedName~OfficeIMO.Tests.Excel.FeatureReport_Preflight_BlocksStructureTemplateAndPdfWhenPreserveOnlyPartsExist'
) -join '|'

$wordAdvancedFilter = @(
    'FullyQualifiedName~OfficeIMO.Tests.Word.LegacyDoc_LoadLegacyDocWithReport_ReportsCompoundStorageAsSupportedMetadata',
    'FullyQualifiedName~OfficeIMO.Tests.Word.LegacyDoc_LoadLegacyDocWithReport_ReportsActiveXAndEmbeddedPackageStorageAsCompoundMetadata',
    'FullyQualifiedName~OfficeIMO.Tests.Word.Test_EmbeddedPayloadCanBeHashedReplacedAndRemoved',
    'FullyQualifiedName~OfficeIMO.Tests.Word.Test_SavingAndRemovingMacros',
    'FullyQualifiedName~OfficeIMO.Tests.Word.FeatureReportPreflight_',
    'FullyQualifiedName=OfficeIMO.Tests.WordImageExportTests.WordDocument_ExportsEveryEstimatedPageAsNamedImagesAndSnapshots',
    'FullyQualifiedName=OfficeIMO.Tests.WordImageExportTests.WordDocument_BatchExportHonorsPageRange',
    'FullyQualifiedName=OfficeIMO.Tests.WordImageExportTests.WordDocument_BatchBuilderSavesSelectedPagesSynchronouslyAndAsynchronously',
    'FullyQualifiedName=OfficeIMO.Tests.WordImageExportTests.WordDocument_BatchExportRejectsInvalidRanges'
) -join '|'

Write-Host "Office interoperability gate suite: $Suite" -ForegroundColor Yellow

if ($Suite -in @('Full', 'Corpus')) {
    Invoke-InteroperabilityGateStep `
        -Name 'Excel corpus manifest identity and load contract' `
        -Project $projects.Excel `
        -Filter 'Category=OfficeInteroperability'

    Invoke-InteroperabilityGateStep `
        -Name 'Word corpus manifest identity and load contract' `
        -Project $projects.Word `
        -Filter 'Category=OfficeInteroperability'

    Invoke-InteroperabilityGateStep `
        -Name 'Legacy XLS approved import and projection-gap reports' `
        -Project $projects.Excel `
        -Filter $excelLegacyCorpusFilter

    Invoke-InteroperabilityGateStep `
        -Name 'Excel-authored XLSB projection, native rewrite, and conversion contract' `
        -Project $projects.Excel `
        -Filter $xlsbExcelAuthoredFilter

    Invoke-InteroperabilityGateStep `
        -Name 'Legacy DOC approved import reports' `
        -Project $projects.Word `
        -Filter 'FullyQualifiedName=OfficeIMO.Tests.Word.LegacyDoc_CorpusImportReports_MatchCheckedInBaselines'
}

if ($Suite -in @('Full', 'Safety')) {
    Invoke-InteroperabilityGateStep `
        -Name 'Shared Word and Excel package security policy' `
        -Project $projects.Drawing `
        -Filter 'FullyQualifiedName~OfficeIMO.Tests.DrawingOfficePackageSecurityTests'

    Invoke-InteroperabilityGateStep `
        -Name 'Excel preservation, review, formula extension, and preflight contract' `
        -Project $projects.Excel `
        -Filter $excelAdvancedFilter

    Invoke-InteroperabilityGateStep `
        -Name 'Word compound payload, macro, preflight, and batch-render contract' `
        -Project $projects.Word `
        -Filter $wordAdvancedFilter
}

Write-Host ""
Write-Host "Office interoperability gate completed for suite: $Suite." -ForegroundColor Green
