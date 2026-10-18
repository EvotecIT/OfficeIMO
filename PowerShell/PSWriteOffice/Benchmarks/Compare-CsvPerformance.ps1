[CmdletBinding()]
param(
    [ValidateSet('Smoke', 'Standard', 'Large', 'Full', 'SuperLarge')]
    [string] $Suite = 'Standard',

    [object[]] $RowCount,

    [int] $RepeatCount = 0,

    [string[]] $Scenario,

    [string[]] $Engine = @('PSWriteOffice', 'NativeCsv'),

    [string] $OutputDirectory = (Join-Path $PSScriptRoot '..\Ignore\Benchmarks\CsvPerformance'),

    [Alias('Plan')]
    [switch] $ListScenarios,

    [string] $OfficeIMORoot,

    [ValidateSet('Debug', 'Release')]
    [string] $PSWriteOfficeConfiguration = 'Release',

    [switch] $SkipPSWriteOfficeBuild,

    [switch] $UpdateReadme,

    [string] $ReadmeBlockId = 'CsvComparison'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..')
$specPath = Join-Path $PSScriptRoot 'Csv\csv-performance.benchmark.ps1'
$projectPath = Join-Path $repoRoot 'Sources\PSWriteOffice\PSWriteOffice.csproj'
$moduleManifest = Join-Path $repoRoot 'PSWriteOffice.psd1'
$benchmarkHelperPath = Join-Path $PSScriptRoot 'Excel\excel-performance.helpers.ps1'
$resultValidationHelperPath = Join-Path $PSScriptRoot 'Benchmark.ResultValidation.ps1'
$officeIMOSourceHelperPath = Join-Path $PSScriptRoot 'OfficeIMO.Source.ps1'

Import-Module PSPublishModule -MinimumVersion 3.0.64 -Force -ErrorAction Stop
if (-not (Get-Command Invoke-BenchmarkSuite -ErrorAction SilentlyContinue)) {
    throw 'The imported PSPublishModule does not expose Invoke-BenchmarkSuite.'
}
. $benchmarkHelperPath
. $resultValidationHelperPath
. $officeIMOSourceHelperPath

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

$expectedRowCounts = @(Resolve-BenchmarkRequestedRowCount -RowCount $RowCount -Suite $Suite)
$selectedCases = @(
    Get-CsvBenchmarkCase -Suite $Suite | Where-Object {
        -not $Scenario -or $Scenario.Count -eq 0 -or @($Scenario) -contains [string]$_.Name
    }
)
$selectedRuns = @(
    foreach ($engineName in @($Engine)) {
        foreach ($case in $selectedCases) {
            if (Test-CsvBenchmarkEngineSupport -Engine $engineName -Case $case) {
                [pscustomobject]@{ Engine = $engineName; Case = $case }
            }
        }
    }
)
$unsupportedRuns = @(
    foreach ($engineName in @($Engine)) {
        foreach ($case in $selectedCases) {
            if (-not (Test-CsvBenchmarkEngineSupport -Engine $engineName -Case $case)) {
                [pscustomobject]@{ Engine = $engineName; Case = $case }
            }
        }
    }
)
if (-not $ListScenarios.IsPresent -and $selectedRuns.Count -eq 0) {
    throw "Requested CSV benchmark lanes are unsupported for engines '$($Engine -join ', ')' and scenarios '$($Scenario -join ', ')'."
}
if (-not $ListScenarios.IsPresent -and
    $PSBoundParameters.ContainsKey('Engine') -and
    $PSBoundParameters.ContainsKey('Scenario') -and
    $unsupportedRuns.Count -gt 0) {
    $unsupportedDescriptions = $unsupportedRuns | ForEach-Object { "$($_.Engine) / $($_.Case.Name)" }
    throw "Requested CSV benchmark lane(s) are unsupported: $($unsupportedDescriptions -join ', ')."
}

$requiresPSWriteOffice = -not $ListScenarios.IsPresent -and (@($Engine) -contains 'PSWriteOffice')
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
        -BlockId $ReadmeBlockId
}
$result
