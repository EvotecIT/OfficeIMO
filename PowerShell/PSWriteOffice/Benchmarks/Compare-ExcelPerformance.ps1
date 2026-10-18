[CmdletBinding()]
param(
    [ValidateSet('Smoke', 'Standard', 'Large', 'Full', 'SuperLarge')]
    [string] $Suite = 'Standard',

    [object[]] $RowCount,

    [int] $RepeatCount = 0,

    [ValidateRange(-1, [int]::MaxValue)]
    [int] $WarmupCount = -1,

    [ValidateRange(-1, [int]::MaxValue)]
    [int] $CooldownMilliseconds = -1,

    [string[]] $Scenario,

    [string[]] $Engine = @('PSWriteOffice', 'ImportExcel', 'ExcelFast'),

    [string] $OutputDirectory = (Join-Path $PSScriptRoot '..\Ignore\Benchmarks\ExcelPerformance'),

    [Alias('Plan')]
    [switch] $ListScenarios,

    [switch] $SkipWorkbookValidation,

    [switch] $SkipImportExcelInstall,

    [switch] $SkipExcelFastInstall,

    [string] $OfficeIMORoot,

    [ValidateSet('Debug', 'Release')]
    [string] $PSWriteOfficeConfiguration = 'Release',

    [switch] $SkipPSWriteOfficeBuild,

    [switch] $UpdateReadme
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..')
$specPath = Join-Path $PSScriptRoot 'Excel\excel-performance.benchmark.ps1'
$projectPath = Join-Path $repoRoot 'Sources\PSWriteOffice\PSWriteOffice.csproj'
$moduleManifest = Join-Path $repoRoot 'PSWriteOffice.psd1'
$benchmarkHelperPath = Join-Path $PSScriptRoot 'Excel\excel-performance.helpers.ps1'
$resultValidationHelperPath = Join-Path $PSScriptRoot 'Benchmark.ResultValidation.ps1'
$officeIMOSourceHelperPath = Join-Path $PSScriptRoot 'OfficeIMO.Source.ps1'

Import-Module PSPublishModule -MinimumVersion 3.0.64 -Force -ErrorAction Stop
if (-not (Get-Command Invoke-BenchmarkSuite -ErrorAction SilentlyContinue)) {
    throw 'The imported PSPublishModule does not expose Invoke-BenchmarkSuite.'
}

$Engine = @(
    foreach ($engineName in @($Engine)) {
        foreach ($part in ([string]$engineName -split ',')) {
            $normalized = $part.Trim()
            if (-not [string]::IsNullOrWhiteSpace($normalized)) {
                $normalized
            }
        }
    }
) | Select-Object -Unique

$Scenario = @(
    foreach ($scenarioName in @($Scenario)) {
        foreach ($part in ([string]$scenarioName -split ',')) {
            $normalized = $part.Trim()
            if (-not [string]::IsNullOrWhiteSpace($normalized)) {
                $normalized
            }
        }
    }
) | Select-Object -Unique

. $benchmarkHelperPath
. $resultValidationHelperPath
. $officeIMOSourceHelperPath
$expectedRowCounts = @(Resolve-BenchmarkRequestedRowCount -RowCount $RowCount -Suite $Suite)
$selectedCases = @(
    Get-ExcelBenchmarkCase -Suite $Suite | Where-Object {
        -not $Scenario -or $Scenario.Count -eq 0 -or @($Scenario) -contains [string]$_.Name
    }
)
$selectedRuns = @(
    foreach ($engineName in @($Engine)) {
        foreach ($case in $selectedCases) {
            if (Test-ExcelBenchmarkEngineCaseSupport -Engine $engineName -Case $case) {
                [pscustomobject]@{ Engine = $engineName; Case = $case }
            }
        }
    }
)
$unsupportedRuns = @(
    foreach ($engineName in @($Engine)) {
        foreach ($case in $selectedCases) {
            if (-not (Test-ExcelBenchmarkEngineCaseSupport -Engine $engineName -Case $case)) {
                [pscustomobject]@{ Engine = $engineName; Case = $case }
            }
        }
    }
)
if (-not $ListScenarios.IsPresent -and $selectedRuns.Count -eq 0) {
    throw "Requested Excel benchmark lanes are unsupported for engines '$($Engine -join ', ')' and scenarios '$($Scenario -join ', ')'."
}
if (-not $ListScenarios.IsPresent -and
    $PSBoundParameters.ContainsKey('Engine') -and
    $PSBoundParameters.ContainsKey('Scenario') -and
    $unsupportedRuns.Count -gt 0) {
    $unsupportedDescriptions = $unsupportedRuns | ForEach-Object { "$($_.Engine) / $($_.Case.Name)" }
    throw "Requested Excel benchmark lane(s) are unsupported: $($unsupportedDescriptions -join ', ')."
}

$requiresExcelFast = -not $ListScenarios.IsPresent -and
    [bool]($selectedRuns | Where-Object Engine -EQ 'ExcelFast' | Select-Object -First 1)
if ($requiresExcelFast) {
    $excelFastModuleRoot = Join-Path $OutputDirectory 'Modules'
    New-Item -ItemType Directory -Force -Path $excelFastModuleRoot | Out-Null
    if (-not (($env:PSModulePath -split [IO.Path]::PathSeparator) -contains $excelFastModuleRoot)) {
        $env:PSModulePath = $excelFastModuleRoot + [IO.Path]::PathSeparator + $env:PSModulePath
    }

    try {
        if (-not (Get-Module -ListAvailable ExcelFast | Sort-Object Version -Descending | Select-Object -First 1)) {
            if ([bool]$SkipExcelFastInstall) {
                throw 'ExcelFast is not installed and -SkipExcelFastInstall was specified.'
            }
            Save-Module -Name ExcelFast -Path $excelFastModuleRoot -Repository PSGallery -AllowPrerelease -Force -ErrorAction Stop
        }
        Import-Module ExcelFast -Force -ErrorAction Stop
    } catch {
        throw "ExcelFast was requested but could not be prepared. The comparison cannot continue without the requested competitor lane. $($_.Exception.Message)"
    }
}

$psWriteOfficeFixtureOperations = @(
    'ReadFullSheet',
    'ReadRange',
    'ReadNoHeaderRange',
    'ReadUsedRangeDataTable',
    'ReadTableMetadata',
    'ReadNamedRangeMetadata'
)
$requiresPSWriteOffice = -not $ListScenarios.IsPresent -and [bool](
    $selectedRuns | Where-Object {
        $_.Engine -eq 'PSWriteOffice' -or
        (-not $SkipWorkbookValidation.IsPresent -and [bool]$_.Case.ValidateWorkbook) -or
        [string]$_.Case.OperationKey -in $psWriteOfficeFixtureOperations
    } | Select-Object -First 1
)
if ($requiresPSWriteOffice) {
    if ([string]::IsNullOrWhiteSpace($OfficeIMORoot)) {
        $OfficeIMORoot = $env:OfficeIMORoot
    }
    if ([string]::IsNullOrWhiteSpace($OfficeIMORoot)) {
        throw 'OfficeIMORoot is required for PSWriteOffice performance comparisons. Benchmarks must build against the current OfficeIMO source tree, not a published package.'
    }

    $OfficeIMORoot = Resolve-OfficeIMOSourceRoot -Path $OfficeIMORoot
    $env:OfficeIMORoot = $OfficeIMORoot
    $env:UseOfficeIMOProjectReferences = 'true'

    if (-not $SkipPSWriteOfficeBuild.IsPresent) {
        & dotnet build $projectPath -c $PSWriteOfficeConfiguration -v:minimal "-p:OfficeIMORoot=$OfficeIMORoot" '-p:UseOfficeIMOProjectReferences=true'
        if ($LASTEXITCODE -ne 0) {
            throw "dotnet build failed for PSWriteOffice ($PSWriteOfficeConfiguration)."
        }
    }

    $env:PSWRITEOFFICE_USE_DEVELOPMENT_BINARIES = 'true'
    $env:PSWRITEOFFICE_DEVELOPMENT_CONFIGURATION = $PSWriteOfficeConfiguration
    Import-Module $moduleManifest -Force -ErrorAction Stop
}

$benchmarkVariables = @{
    Suite = $Suite
    RowCount = if ($RowCount) { ($RowCount -join ',') } else { $null }
    SkipWorkbookValidation = [string]$SkipWorkbookValidation.IsPresent
    SkipImportExcelInstall = [string]$SkipImportExcelInstall.IsPresent
    SkipExcelFastInstall = [string]$SkipExcelFastInstall.IsPresent
}

$invokeSplat = @{
    Path = $specPath
    OutputRoot = $OutputDirectory
    RunMode = $Suite
    Engine = $Engine
    Variable = $benchmarkVariables
}
if ($RepeatCount -gt 0) {
    $invokeSplat.IterationCount = $RepeatCount
}
if ($WarmupCount -ge 0) {
    $invokeSplat.WarmupCount = $WarmupCount
}
if ($CooldownMilliseconds -ge 0) {
    $invokeSplat.CooldownMilliseconds = $CooldownMilliseconds
}
if ($Scenario -and $Scenario.Count -gt 0) {
    $invokeSplat.Case = $Scenario
}
if ($ListScenarios.IsPresent) {
    $invokeSplat.Plan = $true
    $plan = Invoke-BenchmarkSuite @invokeSplat
    $planView = foreach ($item in $plan) {
        [pscustomobject]@{
            Scenario = $item.Scenario
            Engine = $item.Engine
            RowCount = $item.Values.RowCount
            Operation = $item.Values.OperationKey
            Skipped = [bool]$item.IsSkipped
        }
    }
    $planView | Sort-Object Scenario, RowCount, Engine | Format-Table -AutoSize
    return
}

$result = Invoke-BenchmarkSuite @invokeSplat
$failedRows = @($result.Summary | Where-Object { $_.FailureCount -gt 0 -or $_.Status -eq 'Failed' })
if ($failedRows.Count -gt 0) {
    $failedDescriptions = $failedRows | ForEach-Object {
        $reasons = if ($_.FailureReasons -and $_.FailureReasons.Keys.Count -gt 0) {
            $_.FailureReasons.Keys -join ' | '
        } else {
            'No failure reason was recorded.'
        }
        "$($_.Engine) $($_.Scenario): $reasons"
    }
    throw "Benchmark run $($result.RunId) had failed lane(s): $($failedDescriptions -join '; ')"
}

Assert-BenchmarkRequestedLaneCompletion `
    -Summary @($result.Summary) `
    -ExpectedRun $selectedRuns `
    -RowCount $expectedRowCounts

$summaryPath = $result.Artifacts['summary.csv']
if ($summaryPath -and $UpdateReadme.IsPresent) {
    & (Join-Path $PSScriptRoot 'Update-PerformanceBenchmarkReadme.ps1') `
        -SummaryPath $summaryPath `
        -ReadmePath (Join-Path $PSScriptRoot 'README.md') `
        -BlockId ExcelComparison
}
$result
