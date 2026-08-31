[CmdletBinding()]
param(
    [ValidateSet('Debug', 'Release')]
    [string] $Configuration = 'Release',

    [ValidateSet('net8.0', 'net10.0')]
    [string] $Framework = 'net8.0',

    [string] $OutputDirectory,

    [switch] $NoRestore,

    [switch] $NoBuild
)

$ErrorActionPreference = 'Stop'
$repositoryRoot = Split-Path -Parent $PSScriptRoot
$project = Join-Path $repositoryRoot 'Build/PdfQualityCorpus/OfficeIMO.PdfQualityCorpus.Tool.csproj'
$fixtureRoot = Join-Path $repositoryRoot 'OfficeIMO.Pdf.Tests/Pdf/Fixtures/Interoperability'
$manifest = Join-Path $fixtureRoot 'corpus-manifest.json'
if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    $OutputDirectory = Join-Path $repositoryRoot 'artifacts/pdf-quality-corpus'
}
$OutputDirectory = [System.IO.Path]::GetFullPath($OutputDirectory)
$jsonReport = Join-Path $OutputDirectory 'scorecard.json'
$markdownReport = Join-Path $OutputDirectory 'scorecard.md'

if (-not $NoRestore) {
    & dotnet restore $project --nologo
    if ($LASTEXITCODE -ne 0) { throw "PDF quality corpus restore failed with exit code $LASTEXITCODE." }
}
if (-not $NoBuild) {
    & dotnet build $project --configuration $Configuration --framework $Framework --no-restore --nologo
    if ($LASTEXITCODE -ne 0) { throw "PDF quality corpus build failed with exit code $LASTEXITCODE." }
}

New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
& dotnet run --project $project --configuration $Configuration --framework $Framework --no-build -- verify-markdown-contract
if ($LASTEXITCODE -ne 0) { throw "PDF quality corpus Markdown contract failed with exit code $LASTEXITCODE." }
& dotnet run --project $project --configuration $Configuration --framework $Framework --no-build -- verify-runner-contracts
if ($LASTEXITCODE -ne 0) { throw "PDF quality corpus runner contracts failed with exit code $LASTEXITCODE." }
& dotnet run --project $project --configuration $Configuration --framework $Framework --no-build -- run `
    --manifest $manifest `
    --root $fixtureRoot `
    --json $jsonReport `
    --markdown $markdownReport `
    --parallelism 4 `
    --timeout-seconds 60 `
    --max-worker-memory-bytes 1073741824 `
    --max-render-pages 4
if ($LASTEXITCODE -ne 0) { throw "PDF quality corpus execution failed with exit code $LASTEXITCODE." }

$report = Get-Content -LiteralPath $jsonReport -Raw | ConvertFrom-Json
if ($report.schemaVersion -ne 1) { throw 'Unexpected PDF quality corpus report schema version.' }
if ([string]::IsNullOrWhiteSpace($report.configuration.manifestSha256) -or $report.configuration.manifestSha256.Length -ne 64) {
    throw 'PDF quality corpus report does not identify the measured manifest by SHA256.'
}
if ($report.configuration.PSObject.Properties.Name -contains 'manifestPath' -or
    $report.configuration.PSObject.Properties.Name -contains 'rootDirectory' -or
    $report.configuration.PSObject.Properties.Name -contains 'jsonReportPath' -or
    $report.configuration.PSObject.Properties.Name -contains 'markdownReportPath') {
    throw 'PDF quality corpus report leaked machine-specific paths.'
}
if ($report.totals.cases -ne 14) { throw "Expected 14 provenance-bound cases, received $($report.totals.cases)." }
if ($report.totals.passed -ne $report.totals.cases -or $report.totals.failed -ne 0 -or $report.totals.timedOut -ne 0) {
    throw 'PDF quality corpus did not complete every manifested case successfully.'
}
if ($report.totals.operationalChecksPassed -ne $report.totals.operationalChecks) {
    throw 'At least one public PDF API stage failed.'
}
if ($report.totals.expectationsPassed -ne $report.totals.expectations) {
    throw 'At least one pinned PDF expectation failed.'
}
if ($report.totals.peakWorkingSetBytes -le 0) {
    throw 'PDF quality corpus did not record a process-isolated peak working set.'
}
$missingProcessEvidence = @($report.cases | Where-Object {
    $_.workerWallClockMilliseconds -le 0 -or $_.peakWorkingSetBytes -le 0
})
if ($missingProcessEvidence.Count -ne 0) {
    throw "PDF quality corpus omitted process evidence for $($missingProcessEvidence.Count) case(s)."
}
$memoryBudgetFailures = @($report.cases | Where-Object {
    $_.peakWorkingSetBytes -gt $report.configuration.maxWorkerMemoryBytes
})
if ($memoryBudgetFailures.Count -ne 0) {
    throw "PDF quality corpus exceeded its worker memory budget for $($memoryBudgetFailures.Count) case(s)."
}
$incompleteMutationPortfolios = @($report.cases | Where-Object { $_.metrics.mutationPlanCount -ne 21 })
if ($incompleteMutationPortfolios.Count -ne 0) {
    throw "PDF quality corpus did not assess all 21 preservation-safe mutation operations for $($incompleteMutationPortfolios.Count) case(s)."
}
$unprovenComplianceClaims = @($report.cases | Where-Object { $_.metrics.claimableComplianceClaimCount -ne 0 })
if ($unprovenComplianceClaims.Count -ne 0) {
    throw "PDF quality corpus treated unvalidated declared compliance as claimable for $($unprovenComplianceClaims.Count) case(s)."
}
if (-not (Test-Path -LiteralPath $markdownReport)) { throw 'PDF quality corpus Markdown report was not created.' }

Write-Host "PDF quality corpus evidence: $OutputDirectory"
