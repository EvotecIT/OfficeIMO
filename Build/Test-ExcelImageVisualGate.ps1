param(
    [string] $Configuration = "Debug",
    [string] $Framework = "net8.0",
    [ValidateSet("Full", "Smoke", "Architecture")]
    [string] $Suite = "Full",
    [switch] $NoRestore,
    [switch] $NoBuild,
    [switch] $UpdateBaselines,
    [switch] $SkipArchitecture
)

$ErrorActionPreference = 'Stop'

$repoRoot = Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')
$excelTestProject = Join-Path $repoRoot 'OfficeIMO.Excel.Tests/OfficeIMO.Excel.Tests.csproj'
$architectureTestProject = Join-Path $repoRoot 'OfficeIMO.Tests/OfficeIMO.Tests.csproj'

if (-not (Test-Path -LiteralPath $excelTestProject)) {
    throw "OfficeIMO Excel test project was not found: $excelTestProject"
}

if (-not (Test-Path -LiteralPath $architectureTestProject)) {
    throw "OfficeIMO aggregate test project was not found: $architectureTestProject"
}

if ($Suite -eq "Architecture" -and $SkipArchitecture) {
    throw "-Suite Architecture cannot be combined with -SkipArchitecture because it would run no checks."
}

function Invoke-VisualGateStep {
    param(
        [Parameter(Mandatory)]
        [string] $Name,

        [Parameter(Mandatory)]
        [string] $Filter,

        [Parameter(Mandatory)]
        [string] $Project
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

function New-TestNameFilter {
    param(
        [Parameter(Mandatory)]
        [string] $ClassName,

        [Parameter(Mandatory)]
        [string[]] $Names
    )

    ($Names | ForEach-Object { "FullyQualifiedName=$ClassName.$_" }) -join "|"
}

$fullGeneratedFilter = 'FullyQualifiedName~ExcelImageExportVisualBaselineTests&FullyQualifiedName~MatchesApprovedBaselines'
$fullApprovedFilter = 'FullyQualifiedName~ExcelImageExportVisualBaselineTests&FullyQualifiedName~AreRenderableAndNonBlank'
$architectureFilter = 'FullyQualifiedName~DrawingArchitectureTests'
$fidelityManifestFilter = 'FullyQualifiedName~ExcelImageExportVisualFidelityGateTests'

$visualBaselineTestClass = 'OfficeIMO.Tests.ExcelImageExportVisualBaselineTests'

$smokeGeneratedFilter = New-TestNameFilter -ClassName $visualBaselineTestClass -Names @(
    'PremiumRangeImageExportMatchesApprovedBaselines',
    'RichTextImageExportMatchesApprovedBaselines',
    'HeaderFooterImageExportMatchesApprovedBaselines',
    'ChartAxisLabelsImageExportMatchesApprovedBaselines',
    'PageLayoutImageExportMatchesApprovedBaselines',
    'ConditionalFormattingImageExportMatchesApprovedBaselines',
    'ExpandedIconSetImageExportMatchesApprovedBaselines',
    'TextSpillImageExportMatchesApprovedBaselines',
    'CommentBodyImageExportMatchesApprovedBaselines',
    'DrawingObjectImageExportMatchesApprovedBaselines',
    'TransformedImageExportMatchesApprovedBaselines'
)

$smokeApprovedFilter = New-TestNameFilter -ClassName $visualBaselineTestClass -Names @(
    'ApprovedPremiumRangeBaselinesAreRenderableAndNonBlank',
    'ApprovedRichTextBaselinesAreRenderableAndNonBlank',
    'ApprovedHeaderFooterImageBaselinesAreRenderableAndNonBlank',
    'ApprovedChartAxisLabelsBaselinesAreRenderableAndNonBlank',
    'ApprovedPageLayoutBaselinesAreRenderableAndNonBlank',
    'ApprovedConditionalFormattingBaselinesAreRenderableAndNonBlank',
    'ApprovedExpandedIconSetBaselinesAreRenderableAndNonBlank',
    'ApprovedTextSpillBaselinesAreRenderableAndNonBlank',
    'ApprovedCommentBodyBaselinesAreRenderableAndNonBlank',
    'ApprovedDrawingObjectBaselinesAreRenderableAndNonBlank',
    'ApprovedTransformedImageBaselinesAreRenderableAndNonBlank'
)

$previousUpdateBaselines = $env:OFFICEIMO_UPDATE_EXCEL_IMAGE_BASELINES
try {
    if ($UpdateBaselines) {
        $env:OFFICEIMO_UPDATE_EXCEL_IMAGE_BASELINES = '1'
        Write-Host "OFFICEIMO_UPDATE_EXCEL_IMAGE_BASELINES=1" -ForegroundColor Yellow
    } else {
        $env:OFFICEIMO_UPDATE_EXCEL_IMAGE_BASELINES = $null
    }

    Write-Host "Excel image visual gate suite: $Suite" -ForegroundColor Yellow

    if ($Suite -eq "Full") {
        Invoke-VisualGateStep `
            -Name 'Excel image generated output matches approved baselines' `
            -Filter $fullGeneratedFilter `
            -Project $excelTestProject

        Invoke-VisualGateStep `
            -Name 'Approved Excel image baselines are renderable and nonblank' `
            -Filter $fullApprovedFilter `
            -Project $excelTestProject

        Invoke-VisualGateStep `
            -Name 'Excel image visual fidelity manifest tracks clean baselines, approximations, and gaps' `
            -Filter $fidelityManifestFilter `
            -Project $excelTestProject
    } elseif ($Suite -eq "Smoke") {
        Invoke-VisualGateStep `
            -Name 'Excel image smoke output matches approved baselines' `
            -Filter $smokeGeneratedFilter `
            -Project $excelTestProject

        Invoke-VisualGateStep `
            -Name 'Approved Excel image smoke baselines are renderable and nonblank' `
            -Filter $smokeApprovedFilter `
            -Project $excelTestProject

        Invoke-VisualGateStep `
            -Name 'Excel image visual fidelity manifest tracks clean baselines, approximations, and gaps' `
            -Filter $fidelityManifestFilter `
            -Project $excelTestProject
    }

    if (-not $SkipArchitecture) {
        Invoke-VisualGateStep `
            -Name 'Shared Drawing image-rendering architecture guard' `
            -Filter $architectureFilter `
            -Project $architectureTestProject
    }
} finally {
    $env:OFFICEIMO_UPDATE_EXCEL_IMAGE_BASELINES = $previousUpdateBaselines
}

Write-Host ""
Write-Host "Excel image visual gate completed for suite: $Suite." -ForegroundColor Green
