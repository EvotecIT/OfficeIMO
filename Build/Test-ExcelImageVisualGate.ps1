param(
    [string] $Configuration = "Debug",
    [string] $Framework = "net8.0",
    [switch] $NoRestore,
    [switch] $NoBuild,
    [switch] $UpdateBaselines,
    [switch] $SkipArchitecture
)

$ErrorActionPreference = 'Stop'

$repoRoot = Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')
$testProject = Join-Path $repoRoot 'OfficeIMO.Tests/OfficeIMO.Tests.csproj'

if (-not (Test-Path -LiteralPath $testProject)) {
    throw "OfficeIMO test project was not found: $testProject"
}

function Invoke-VisualGateStep {
    param(
        [Parameter(Mandatory)]
        [string] $Name,

        [Parameter(Mandatory)]
        [string] $Filter
    )

    Write-Host ""
    Write-Host "== $Name ==" -ForegroundColor Cyan
    Write-Host "Filter: $Filter" -ForegroundColor DarkCyan

    $arguments = @(
        'test',
        $testProject,
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

$previousUpdateBaselines = $env:OFFICEIMO_UPDATE_EXCEL_IMAGE_BASELINES
try {
    if ($UpdateBaselines) {
        $env:OFFICEIMO_UPDATE_EXCEL_IMAGE_BASELINES = '1'
        Write-Host "OFFICEIMO_UPDATE_EXCEL_IMAGE_BASELINES=1" -ForegroundColor Yellow
    } else {
        $env:OFFICEIMO_UPDATE_EXCEL_IMAGE_BASELINES = $null
    }

    Invoke-VisualGateStep `
        -Name 'Excel image generated output matches approved baselines' `
        -Filter 'FullyQualifiedName~ExcelImageExportVisualBaselineTests&FullyQualifiedName~MatchesApprovedBaselines'

    Invoke-VisualGateStep `
        -Name 'Approved Excel image baselines are renderable and nonblank' `
        -Filter 'FullyQualifiedName~ExcelImageExportVisualBaselineTests&FullyQualifiedName~AreRenderableAndNonBlank'

    if (-not $SkipArchitecture) {
        Invoke-VisualGateStep `
            -Name 'Shared Drawing image-rendering architecture guard' `
            -Filter 'FullyQualifiedName~DrawingArchitectureTests'
    }
} finally {
    $env:OFFICEIMO_UPDATE_EXCEL_IMAGE_BASELINES = $previousUpdateBaselines
}

Write-Host ""
Write-Host "Excel image visual gate completed." -ForegroundColor Green
