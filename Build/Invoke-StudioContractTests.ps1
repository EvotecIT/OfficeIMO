[CmdletBinding()]
param(
    [string] $Configuration = 'Release',
    [string] $ResultsDirectory = 'TestResults/Studio',
    [string] $TrxPath = '',
    [ValidateSet('Auto', 'Windows', 'Other')]
    [string] $Platform = 'Auto'
)

$ErrorActionPreference = 'Stop'
$PSNativeCommandUseErrorActionPreference = $false

$repositoryRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
$resolvedResultsDirectory = if ([System.IO.Path]::IsPathRooted($ResultsDirectory)) {
    $ResultsDirectory
} else {
    Join-Path $repositoryRoot $ResultsDirectory
}

$testExitCode = 1
if ([string]::IsNullOrWhiteSpace($TrxPath)) {
    & dotnet test (Join-Path $repositoryRoot 'OfficeIMO.Studio.Tests\OfficeIMO.Studio.Tests.csproj') `
        --configuration $Configuration `
        --logger 'trx;LogFileName=studio.trx' `
        --results-directory $resolvedResultsDirectory
    $testExitCode = $LASTEXITCODE
    if ($testExitCode -eq 0) {
        exit 0
    }
    if ($testExitCode -ne 1) {
        exit $testExitCode
    }

    $TrxPath = Join-Path $resolvedResultsDirectory 'studio.trx'
}

$runningOnWindows = switch ($Platform) {
    'Windows' { $true }
    'Other' { $false }
    default {
        if (-not [string]::IsNullOrWhiteSpace($env:RUNNER_OS)) {
            $env:RUNNER_OS -eq 'Windows'
        } else {
            [System.OperatingSystem]::IsWindows()
        }
    }
}

if (-not $runningOnWindows -or -not (Test-Path -LiteralPath $TrxPath -PathType Leaf)) {
    exit $testExitCode
}

[xml] $trx = Get-Content -LiteralPath $TrxPath -Raw
$namespace = [System.Xml.XmlNamespaceManager]::new($trx.NameTable)
$namespace.AddNamespace('t', 'http://microsoft.com/schemas/VisualStudio/TeamTest/2010')
$failedResults = @($trx.SelectNodes('//t:UnitTestResult[@outcome="Failed"]', $namespace))
if ($failedResults.Count -eq 0) {
    exit $testExitCode
}

$resultSummaries = @($trx.SelectNodes('/t:TestRun/t:ResultSummary', $namespace))
$resultSummary = if ($resultSummaries.Count -eq 1) { $resultSummaries[0] } else { $null }
$counters = if ($null -ne $resultSummary) { $resultSummary.SelectSingleNode('t:Counters', $namespace) } else { $null }
if ($resultSummaries.Count -ne 1 -or
    [string] $resultSummary.GetAttribute('outcome') -ne 'Failed' -or
    $null -eq $counters) {
    Write-Error 'Studio test suppression requires one failed ResultSummary with counters.'
    exit $testExitCode
}

$counterNames = @(
    'total', 'executed', 'passed', 'failed', 'error', 'timeout', 'aborted',
    'inconclusive', 'passedButRunAborted', 'notRunnable', 'notExecuted',
    'disconnected', 'warning', 'completed', 'inProgress', 'pending'
)
$counterValues = @{}
foreach ($counterName in $counterNames) {
    $attribute = $counters.GetAttributeNode($counterName)
    [long] $counterValue = 0
    if ($null -eq $attribute -or
        -not [long]::TryParse($attribute.Value, [System.Globalization.NumberStyles]::None,
            [System.Globalization.CultureInfo]::InvariantCulture, [ref] $counterValue) -or
        $counterValue -lt 0) {
        Write-Error "Studio test suppression requires a valid '$counterName' counter."
        exit $testExitCode
    }
    $counterValues[$counterName] = $counterValue
}

$allResults = @($trx.SelectNodes('//t:UnitTestResult', $namespace))
$exceptionalCounterNames = @(
    'error', 'timeout', 'aborted', 'inconclusive', 'passedButRunAborted',
    'notRunnable', 'notExecuted', 'disconnected', 'warning', 'inProgress', 'pending'
)
$hasExceptionalCounter = @($exceptionalCounterNames | Where-Object { $counterValues[$_] -ne 0 }).Count -gt 0
$runLevelProblems = @($resultSummary.SelectNodes('.//t:TestRunMessage | .//t:ErrorInfo', $namespace))
$failedTestNames = @($failedResults | ForEach-Object { [string] $_.GetAttribute('testName') })
foreach ($runInfo in @($resultSummary.SelectNodes('.//t:RunInfo', $namespace))) {
    $runInfoText = [string] $runInfo.Text
    $describesRecordedFailure = [string] $runInfo.GetAttribute('outcome') -eq 'Error' -and
        $runInfoText.IndexOf('[FAIL]', [System.StringComparison]::Ordinal) -ge 0 -and
        @($failedTestNames | Where-Object {
            $_.Length -gt 0 -and $runInfoText.IndexOf($_, [System.StringComparison]::Ordinal) -ge 0
        }).Count -eq 1
    if (-not $describesRecordedFailure) { $runLevelProblems += $runInfo }
}
$runIsComplete = $counterValues.total -gt 0 -and
    $counterValues.executed -eq $counterValues.total -and
    ($counterValues.completed -eq 0 -or $counterValues.completed -eq $counterValues.total) -and
    ($counterValues.passed + $counterValues.failed) -eq $counterValues.total -and
    $counterValues.failed -eq $failedResults.Count -and
    $allResults.Count -eq $counterValues.total
if (-not $runIsComplete -or $hasExceptionalCounter -or $runLevelProblems.Count -gt 0) {
    Write-Error 'Studio tests were incomplete or contain run-level failures; the Avalonia teardown exemption was not applied.'
    exit $testExitCode
}

$knownMessage = 'System.NullReferenceException : Object reference not set to an instance of an object.'
$knownStackPrefix = 'at Avalonia.Headless.HeadlessUnitTestSession.Dispose()'
$unexpectedFailures = @($failedResults | Where-Object {
    $message = [string] $_.Output.ErrorInfo.Message
    $stackTrace = ([string] $_.Output.ErrorInfo.StackTrace).TrimStart()
    $message.Trim() -ne $knownMessage -or -not $stackTrace.StartsWith($knownStackPrefix, [System.StringComparison]::Ordinal)
})

if ($unexpectedFailures.Count -gt 0) {
    Write-Error "Studio tests contain $($unexpectedFailures.Count) failure(s) outside the temporary Avalonia teardown exemption."
    exit $testExitCode
}

Write-Warning "Suppressing $($failedResults.Count) Windows-only Avalonia.Headless session disposal failure(s). Upstream: https://github.com/AvaloniaUI/Avalonia/pull/22222"
exit 0
