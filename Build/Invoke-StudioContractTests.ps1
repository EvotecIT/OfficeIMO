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
