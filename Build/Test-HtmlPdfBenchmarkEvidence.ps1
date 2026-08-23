param(
    [string] $EvidenceRoot = (Join-Path $PSScriptRoot '../Website/static/data/benchmarks/library-comparisons')
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repositoryRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
$resolvedEvidenceRoot = (Resolve-Path -LiteralPath $EvidenceRoot).Path
$indexPath = Join-Path $resolvedEvidenceRoot 'index.json'
$index = Get-Content -LiteralPath $indexPath -Raw | ConvertFrom-Json
$comparisonId = 'pdf-html-generation-net10.0'
$expectedScenarios = @('Chromium', 'ITextPdfHtml', 'OfficeIMO', 'PeachPDF')
$expectedScales = @('Easy', 'Medium', 'High')
$sourceCommit = $null
$browserCommit = $null

foreach ($platform in @('windows', 'linux')) {
    $fileName = "$comparisonId-$platform-full.json"
    $resultPath = Join-Path $resolvedEvidenceRoot $fileName
    $result = Get-Content -LiteralPath $resultPath -Raw | ConvertFrom-Json
    $entry = @($index.entries | Where-Object {
            $_.comparisonId -eq $comparisonId -and
            $_.platform -eq $platform -and
            $_.runMode -eq 'full'
        })

    if ($entry.Count -ne 1 -or $entry[0].publish -ne $true -or $entry[0].comparable -ne $true) {
        throw "HTML/PDF benchmark evidence must contain one publishable, comparable $platform full-run catalog entry."
    }

    $actualHash = (Get-FileHash -LiteralPath $resultPath -Algorithm SHA256).Hash.ToLowerInvariant()
    if (-not [string]::Equals($actualHash, [string] $entry[0].resultSha256, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "HTML/PDF benchmark evidence hash does not match the catalog for $fileName."
    }

    if ([string] $result.suite -ne 'OfficeIMO.Pdf.HtmlGeneration' -or
        [string] $result.environment.osFamily -ne (Get-Culture).TextInfo.ToTitleCase($platform) -or
        [string] $result.environment.runner -ne 'BenchmarkDotNet' -or
        [string] $result.metadata.'benchmark.workload.framework' -ne 'net10.0' -or
        [string] $result.metadata.runMode -ne 'full' -or
        [string] $result.metadata.gitWorktreeClean -ne 'true') {
        throw "HTML/PDF benchmark evidence has invalid environment or provenance metadata for $fileName."
    }

    $currentSourceCommit = [string] $result.metadata.'benchmark.workload.sourceCommit'
    $currentBrowserCommit = [string] $result.metadata.'benchmark.browser.sourceCommit'
    if ($currentSourceCommit -notmatch '^[0-9a-f]{40}$' -or
        $currentBrowserCommit -notmatch '^[0-9a-f]{40}$' -or
        -not [string]::Equals($currentSourceCommit, [string] $result.metadata.gitSha, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "HTML/PDF benchmark evidence is not bound to exact OfficeIMO and HtmlTinkerX source commits for $fileName."
    }

    if ($null -eq $sourceCommit) {
        $sourceCommit = $currentSourceCommit
        $browserCommit = $currentBrowserCommit
    } elseif (-not [string]::Equals($sourceCommit, $currentSourceCommit, [System.StringComparison]::OrdinalIgnoreCase) -or
        -not [string]::Equals($browserCommit, $currentBrowserCommit, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw 'Windows and Linux HTML/PDF benchmark evidence must measure the same OfficeIMO and HtmlTinkerX source commits.'
    }

    $summary = @($result.summary)
    if ($summary.Count -ne ($expectedScenarios.Count * $expectedScales.Count)) {
        throw "HTML/PDF benchmark evidence must contain all 12 engine/scale cases for $fileName."
    }

    foreach ($scenario in $expectedScenarios) {
        foreach ($scale in $expectedScales) {
            $cases = @($summary | Where-Object {
                    $_.scenario -eq $scenario -and $_.variables.Scale -eq $scale
                })
            if ($cases.Count -ne 1 -or
                [string] $cases[0].status -ne 'Succeeded' -or
                [int] $cases[0].failureCount -ne 0 -or
                [int] $cases[0].sampleCount -lt 1 -or
                [double] $cases[0].meanMs -le 0 -or
                [double] $cases[0].metrics.BytesAllocatedPerOperation -le 0) {
                throw "HTML/PDF benchmark case $scenario/$scale is missing or unsuccessful for $fileName."
            }
        }
    }
}

& git -C $repositoryRoot cat-file -e "$sourceCommit`^{commit}"
if ($LASTEXITCODE -ne 0) {
    throw "Measured OfficeIMO source commit $sourceCommit is not available in this repository."
}

& git -C $repositoryRoot merge-base --is-ancestor $sourceCommit HEAD
if ($LASTEXITCODE -ne 0) {
    throw "Measured OfficeIMO source commit $sourceCommit is not an ancestor of the release candidate."
}

$measuredPaths = @(
    'OfficeIMO.Html.Pdf',
    # The public HtmlTinkerX package pin is proven by Test-HtmlPdfBrowserPackages.ps1.
    # Benchmarks use the exact HtmlTinkerX source checkout, so a pin-only project-file
    # change must not invalidate otherwise identical measured runtime code.
    ':(glob)OfficeIMO.Html.Pdf.Browser/**/*.cs',
    'OfficeIMO.Pdf',
    'OfficeIMO.Pdf.Benchmarks',
    'OfficeIMO.Pdf.Benchmarks.Comparisons',
    'Build/LibraryComparisonRunner',
    'Build/Run-LibraryComparisonBenchmarks.ps1',
    'Build/Test-LibraryComparisonRunnerContract.ps1'
)
& git -C $repositoryRoot diff --quiet $sourceCommit HEAD -- @measuredPaths
if ($LASTEXITCODE -ne 0) {
    throw 'HTML/PDF benchmark evidence is stale: measured production or benchmark sources changed after the recorded source commit.'
}

& git -C $repositoryRoot diff --quiet -- @measuredPaths
if ($LASTEXITCODE -ne 0) {
    throw 'HTML/PDF benchmark evidence is stale: measured production or benchmark sources have uncommitted changes.'
}

& git -C $repositoryRoot diff --cached --quiet -- @measuredPaths
if ($LASTEXITCODE -ne 0) {
    throw 'HTML/PDF benchmark evidence is stale: measured production or benchmark sources have staged changes.'
}

Write-Host "Current-source Windows/Linux HTML/PDF benchmark evidence verified at OfficeIMO $sourceCommit and HtmlTinkerX $browserCommit."
