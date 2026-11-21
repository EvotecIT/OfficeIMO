param(
    [string] $EvidenceRoot = (Join-Path $PSScriptRoot '../Website/static/data/benchmarks/evidence')
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repositoryRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
$resolvedEvidenceRoot = (Resolve-Path -LiteralPath $EvidenceRoot).Path
$indexPath = Join-Path $resolvedEvidenceRoot 'index.json'
$index = Get-Content -LiteralPath $indexPath -Raw | ConvertFrom-Json
$comparisonId = 'pdf-html-generation'
$declaration = @($index.comparisons | Where-Object comparisonId -eq $comparisonId)
if ($declaration.Count -ne 1) {
    throw "The benchmark catalog does not declare '$comparisonId'."
}
$expectedLibraries = @($declaration[0].libraries)
$expectedScales = @('Easy', 'Medium', 'High')
$pdfPeckerVersion = [string] @($index.libraries | Where-Object id -eq 'PdfPecker')[0].version
$sourceCommit = $null
$browserCommit = $null
$pdfPeckerPackageSha512 = $null

foreach ($platform in @('windows', 'linux')) {
    $label = "$comparisonId on $platform"
    $entry = @($index.entries | Where-Object {
            $_.comparisonId -eq $comparisonId -and
            $_.platform -eq $platform -and
            $_.framework -eq 'net10.0' -and
            $_.profile -eq 'full' -and
            @($_.parameters.PSObject.Properties).Count -eq 0
        })

    if ($entry.Count -ne 1 -or $entry[0].publish -ne $true -or $entry[0].comparable -ne $true) {
        throw "HTML/PDF benchmark evidence must contain one published, comparable net10.0 full run of $label."
    }

    # Follow the immutable artifact selected by the catalog after a bundle refresh.
    $resultPath = Join-Path $resolvedEvidenceRoot $entry[0].artifactRelativePath
    $result = Get-Content -LiteralPath $resultPath -Raw | ConvertFrom-Json
    $actualHash = (Get-FileHash -LiteralPath $resultPath -Algorithm SHA256).Hash.ToLowerInvariant()
    if (-not [string]::Equals($actualHash, [string] $entry[0].resultSha256, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "HTML/PDF benchmark evidence hash does not match the catalog for $label."
    }

    if ([string] $result.suite -ne 'OfficeIMO.Pdf.HtmlGeneration' -or
        [string] $result.environment.osFamily -ne (Get-Culture).TextInfo.ToTitleCase($platform) -or
        [string] $result.environment.runner -ne 'BenchmarkDotNet' -or
        [string] $result.metadata.'benchmark.workload.framework' -ne 'net10.0' -or
        [string] $result.metadata.'benchmark.package.pdfpecker' -cne $pdfPeckerVersion -or
        [string] $result.metadata.'benchmark.dependency.pdfpecker.sourceCommit' -notmatch '^[0-9a-f]{40}$' -or
        [string] $result.metadata.'benchmark.dependency.pdfpecker.licenseSha256' -notmatch '^[0-9a-f]{64}$' -or
        [string] $result.metadata.'benchmark.dependency.pdfpecker.thirdPartyNoticesSha256' -notmatch '^[0-9a-f]{64}$' -or
        [string] $result.metadata.'benchmark.dependency.pdfpecker.packageSha512' -notmatch '^[A-Za-z0-9+/]{86}==$' -or
        [string] $result.metadata.gitWorktreeClean -ne 'true') {
        throw "HTML/PDF benchmark evidence has invalid environment or provenance metadata for $label."
    }

    $currentPackageSha512 = [string] $result.metadata.'benchmark.dependency.pdfpecker.packageSha512'
    if ($null -eq $pdfPeckerPackageSha512) {
        $pdfPeckerPackageSha512 = $currentPackageSha512
    } elseif (-not [string]::Equals(
        $pdfPeckerPackageSha512, $currentPackageSha512, [System.StringComparison]::Ordinal)) {
        throw 'Windows and Linux HTML/PDF benchmark evidence must use the same PdfPecker package archive SHA-512.'
    }

    $currentSourceCommit = [string] $result.metadata.'benchmark.workload.sourceCommit'
    $currentBrowserCommit = [string] $result.metadata.'benchmark.browser.sourceCommit'
    if ($currentSourceCommit -notmatch '^[0-9a-f]{40}$' -or
        $currentBrowserCommit -notmatch '^[0-9a-f]{40}$' -or
        -not [string]::Equals($currentSourceCommit, [string] $result.metadata.gitSha, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "HTML/PDF benchmark evidence is not bound to exact OfficeIMO and HtmlTinkerX source commits for $label."
    }

    if ($null -eq $sourceCommit) {
        $sourceCommit = $currentSourceCommit
        $browserCommit = $currentBrowserCommit
    } elseif (-not [string]::Equals($sourceCommit, $currentSourceCommit, [System.StringComparison]::OrdinalIgnoreCase) -or
        -not [string]::Equals($browserCommit, $currentBrowserCommit, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw 'Windows and Linux HTML/PDF benchmark evidence must measure the same OfficeIMO and HtmlTinkerX source commits.'
    }

    # Every engine at every scale has a published cell: measured, validated and within the publish policy.
    foreach ($library in $expectedLibraries) {
        foreach ($scale in $expectedScales) {
            $cells = @($entry[0].cells | Where-Object { $_.library -eq $library -and $_.parameters.Scale -eq $scale })
            if ($cells.Count -ne 1 -or [string] $cells[0].status -ne 'Succeeded' -or $cells[0].publishable -ne $true) {
                throw "HTML/PDF benchmark cell $library/$scale is missing or not published for $label."
            }
        }
    }
}

& git -C $repositoryRoot cat-file -e "$sourceCommit`^{commit}"
if ($LASTEXITCODE -ne 0) {
    throw "Measured OfficeIMO source commit $sourceCommit is not available in this repository."
}

& git -C $repositoryRoot merge-base --is-ancestor $sourceCommit HEAD
$measuredCommitIsAncestor = $LASTEXITCODE -eq 0

$measuredPaths = @(
    'OfficeIMO.Core',
    'OfficeIMO.Html',
    'OfficeIMO.Html.Core',
    'OfficeIMO.Html.AngleSharp',
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
    'Build/benchmark-declarations.json',
    'Build/Test-LibraryComparisonRunnerContract.ps1'
)
& git -C $repositoryRoot diff --quiet $sourceCommit HEAD -- @measuredPaths
if ($LASTEXITCODE -ne 0) {
    $relationship = $measuredCommitIsAncestor ? 'after the recorded source commit' : 'relative to the non-ancestor recorded source tree'
    throw "HTML/PDF benchmark evidence is stale: measured production or benchmark sources changed $relationship."
}
if (-not $measuredCommitIsAncestor) {
    Write-Verbose "Measured OfficeIMO benchmark commit $sourceCommit is squash-equivalent to HEAD for every measured path."
}

& git -C $repositoryRoot diff --quiet -- @measuredPaths
if ($LASTEXITCODE -ne 0) {
    throw 'HTML/PDF benchmark evidence is stale: measured production or benchmark sources have uncommitted changes.'
}

& git -C $repositoryRoot diff --cached --quiet -- @measuredPaths
if ($LASTEXITCODE -ne 0) {
    throw 'HTML/PDF benchmark evidence is stale: measured production or benchmark sources have staged changes.'
}
$untrackedMeasuredPaths = @(& git -C $repositoryRoot ls-files --others --exclude-standard -- @measuredPaths)
if ($LASTEXITCODE -ne 0) {
    throw 'Unable to inspect untracked HTML/PDF benchmark source files.'
}
if ($untrackedMeasuredPaths.Count -ne 0) {
    throw "HTML/PDF benchmark evidence is stale: measured sources contain untracked files: $($untrackedMeasuredPaths -join ', ')"
}

Write-Host "Current-source Windows/Linux HTML/PDF benchmark evidence verified at OfficeIMO $sourceCommit and HtmlTinkerX $browserCommit."
