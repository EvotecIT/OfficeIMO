param(
    [string] $Configuration = "Debug",
    [string] $Framework = "net8.0",
    [ValidateSet("Full", "Corpus", "Safety")]
    [string] $Suite = "Full",
    [switch] $NoRestore,
    [switch] $NoBuild,
    [switch] $MicrosoftOffice
)

$ErrorActionPreference = 'Stop'
$PSNativeCommandUseErrorActionPreference = $true

$repoRoot = Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')
$projects = @{
    CompatibilityCatalog = Join-Path $repoRoot 'Build/CompatibilityCatalog/OfficeIMO.CompatibilityCatalog.Tool.csproj'
    Drawing = Join-Path $repoRoot 'OfficeIMO.Drawing.Tests/OfficeIMO.Drawing.Tests.csproj'
    Excel = Join-Path $repoRoot 'OfficeIMO.Excel.Tests/OfficeIMO.Excel.Tests.csproj'
    PowerPoint = Join-Path $repoRoot 'OfficeIMO.PowerPoint.Tests/OfficeIMO.PowerPoint.Tests.csproj'
    Word = Join-Path $repoRoot 'OfficeIMO.Word.Tests/OfficeIMO.Word.Tests.csproj'
}

function Test-CompatibilityCatalogArtifacts {
    Write-Host ""
    Write-Host "== Generated Office format and capability contracts ==" -ForegroundColor Cyan
    $arguments = @(
        'run',
        '--project', $projects.CompatibilityCatalog,
        '--configuration', $Configuration,
        '--framework', $Framework
    )
    if ($NoRestore) {
        $arguments += '--no-restore'
    }
    if ($NoBuild) {
        $arguments += '--no-build'
    }
    $arguments += @('--', '--output', (Join-Path $repoRoot 'Docs/Compatibility/generated'), '--verify')

    Push-Location $repoRoot
    try {
        & dotnet @arguments
        if ($LASTEXITCODE -ne 0) {
            throw "Generated Office compatibility contract verification failed with exit code $LASTEXITCODE."
        }
    } finally {
        Pop-Location
    }
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

Test-CompatibilityCatalogArtifacts

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
        -Name 'PowerPoint corpus identity, preflight, conversion, and reopen contract' `
        -Project $projects.PowerPoint `
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

    Invoke-InteroperabilityGateStep `
        -Name 'PowerPoint binary visual baselines' `
        -Project $projects.PowerPoint `
        -Filter 'FullyQualifiedName=OfficeIMO.Tests.PowerPointLegacyPptTests.ShapeImport_MatchesLibreOfficeVisualReferenceWithinDocumentedTolerance|FullyQualifiedName=OfficeIMO.Tests.PowerPointLegacyPptTests.MicrosoftAuthoredImport_DoesNotRenderMasterPlaceholderPrompts'
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

if ($MicrosoftOffice) {
    if (-not [Runtime.InteropServices.RuntimeInformation]::IsOSPlatform(
        [Runtime.InteropServices.OSPlatform]::Windows)) {
        throw 'The Microsoft Office interoperability lane requires Windows and desktop Word, Excel, and PowerPoint.'
    }

    $previousDocCom = $env:OFFICEIMO_RUN_LEGACY_DOC_COM_VALIDATION
    $previousXlsCom = $env:OFFICEIMO_RUN_LEGACY_XLS_COM_VALIDATION
    $previousPptCom = $env:OFFICEIMO_RUN_LEGACY_PPT_COM_VALIDATION
    try {
        $env:OFFICEIMO_RUN_LEGACY_DOC_COM_VALIDATION = '1'
        $env:OFFICEIMO_RUN_LEGACY_XLS_COM_VALIDATION = '1'
        $env:OFFICEIMO_RUN_LEGACY_PPT_COM_VALIDATION = '1'

        Invoke-InteroperabilityGateStep `
            -Name 'Microsoft Word desktop source and generated conversion oracle' `
            -Project $projects.Word `
            -Filter 'Category=MicrosoftOfficeInteroperability'

        Invoke-InteroperabilityGateStep `
            -Name 'Microsoft Excel desktop corpus source and conversion oracle' `
            -Project $projects.Excel `
            -Filter 'Category=MicrosoftOfficeInteroperability'

        Invoke-InteroperabilityGateStep `
            -Name 'Microsoft PowerPoint desktop corpus source and conversion oracle' `
            -Project $projects.PowerPoint `
            -Filter 'Category=MicrosoftOfficeInteroperability'
    } finally {
        $env:OFFICEIMO_RUN_LEGACY_DOC_COM_VALIDATION = $previousDocCom
        $env:OFFICEIMO_RUN_LEGACY_XLS_COM_VALIDATION = $previousXlsCom
        $env:OFFICEIMO_RUN_LEGACY_PPT_COM_VALIDATION = $previousPptCom
    }
}

Write-Host ""
Write-Host "Office interoperability gate completed for suite: $Suite." -ForegroundColor Green
