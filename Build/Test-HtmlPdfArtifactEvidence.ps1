param(
    [string] $EvidenceRoot = (Join-Path $PSScriptRoot '../Docs/benchmarks/html-pdf-artifact-evidence')
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repositoryRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
$expectedEvidenceRoot = [System.IO.Path]::GetFullPath($EvidenceRoot)
if (-not (Test-Path -LiteralPath $expectedEvidenceRoot -PathType Container)) {
    throw "HTML/PDF artifact evidence directory is missing: $expectedEvidenceRoot"
}
$resolvedEvidenceRoot = (Resolve-Path -LiteralPath $expectedEvidenceRoot).Path
$expectedEngines = @('OfficeIMO', 'PeachPDF', 'ITextPdfHtml', 'Chromium')
$expectedCancellation = @{
    OfficeIMO = @{ Supports = $true; Status = 'Passed' }
    Chromium = @{ Supports = $true; Status = 'Passed' }
    PeachPDF = @{ Supports = $false; Status = 'Unsupported' }
    ITextPdfHtml = @{ Supports = $false; Status = 'Unsupported' }
}
$officeCommit = $null
$browserCommit = $null
$inputHash = $null
$inputSize = $null
$commonTextLength = $null

foreach ($platform in @('windows', 'linux')) {
    $path = Join-Path $resolvedEvidenceRoot "html-pdf-artifact-evidence-net10.0-$platform-high.json"
    $summary = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
    $report = $summary.report
    if ($summary.schemaVersion -ne 1 -or
        [string] $summary.format -ne 'officeimo.html-pdf-artifact-evidence-summary' -or
        [string] $summary.platform -ne $platform -or
        [int] $summary.artifactBundle.artifactCount -ne 37 -or
        [long] $summary.artifactBundle.totalBytes -le 0 -or
        [string] $summary.artifactBundle.manifestSha256 -notmatch '^[0-9a-f]{64}$') {
        throw "Committed $platform HTML/PDF artifact evidence has an invalid summary contract."
    }

    $manifest = @($summary.artifactBundle.artifacts | Sort-Object relativePath)
    if ($manifest.Count -ne 37 -or
        @($manifest | Group-Object { $_.relativePath } | Where-Object { $_.Count -ne 1 }).Count -ne 0 -or
        @($manifest | Where-Object { $_.sizeBytes -le 0 -or $_.sha256 -notmatch '^[0-9a-f]{64}$' }).Count -ne 0) {
        throw "Committed $platform HTML/PDF artifact manifest is incomplete."
    }
    $manifestText = ($manifest | ForEach-Object {
            "$($_.kind)|$($_.relativePath)|$($_.sizeBytes)|$($_.sha256)"
        }) -join "`n"
    $manifestHash = [Convert]::ToHexString(
        [System.Security.Cryptography.SHA256]::HashData(
            [System.Text.Encoding]::UTF8.GetBytes($manifestText))).ToLowerInvariant()
    if (-not [string]::Equals(
        $manifestHash,
        [string] $summary.artifactBundle.manifestSha256,
        [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Committed $platform HTML/PDF artifact manifest hash is invalid."
    }

    $expectedOs = $platform -eq 'windows' ? 'windows' : 'linux|ubuntu'
    if ($report.schemaVersion -ne 2 -or
        [string] $report.scale -ne 'High' -or
        [int] $report.iterations -ne 3 -or
        [string] $report.environment.osDescription -notmatch "(?i)$expectedOs" -or
        [string] $report.environment.externalRasterizer -notmatch '^pdftoppm version ' -or
        $report.provenance.officeIMO.worktreeClean -ne $true -or
        $report.provenance.htmlTinkerX.worktreeClean -ne $true) {
        throw "Committed $platform HTML/PDF artifact report has invalid scale, environment, or provenance."
    }

    $currentOfficeCommit = [string] $report.provenance.officeIMO.commit
    $currentBrowserCommit = [string] $report.provenance.htmlTinkerX.commit
    if ($currentOfficeCommit -notmatch '^[0-9a-f]{40}$' -or $currentBrowserCommit -notmatch '^[0-9a-f]{40}$') {
        throw "Committed $platform HTML/PDF artifact report lacks exact source commits."
    }
    if ($null -eq $officeCommit) {
        $officeCommit = $currentOfficeCommit
        $browserCommit = $currentBrowserCommit
        $inputHash = [string] $report.input.sha256
        $inputSize = [long] $report.input.sizeBytes
    } elseif (-not [string]::Equals($officeCommit, $currentOfficeCommit, [System.StringComparison]::OrdinalIgnoreCase) -or
        -not [string]::Equals($browserCommit, $currentBrowserCommit, [System.StringComparison]::OrdinalIgnoreCase) -or
        -not [string]::Equals($inputHash, [string] $report.input.sha256, [System.StringComparison]::OrdinalIgnoreCase) -or
        $inputSize -ne [long] $report.input.sizeBytes) {
        throw 'Windows and Linux HTML/PDF artifact evidence must use identical source commits and input bytes.'
    }

    $engines = @($report.engines)
    $actualEngineSet = @($engines.engine | Sort-Object) -join '|'
    $expectedEngineSet = @($expectedEngines | Sort-Object) -join '|'
    if ($engines.Count -ne $expectedEngines.Count -or
        $actualEngineSet -ne $expectedEngineSet) {
        throw "Committed $platform HTML/PDF artifact evidence does not cover all four required engines."
    }

    foreach ($engineName in $expectedEngines) {
        $engine = @($engines | Where-Object engine -eq $engineName)[0]
        $cancellation = $expectedCancellation[$engineName]
        if ($engine.memoryComparable -ne $true -or
            $engine.cancellation.apiSupportsCancellation -ne $cancellation.Supports -or
            [string] $engine.cancellation.status -ne $cancellation.Status -or
            $engine.determinism.semanticOutputIdentical -ne $true -or
            $engine.determinism.managedVisualPreviewIdentical -ne $true -or
            $engine.determinism.externalVisualPreviewIdentical -ne $true) {
            throw "Committed $platform $engineName evidence failed cancellation, memory, semantic, or visual gates."
        }
        if ($engineName -eq 'OfficeIMO' -and $engine.determinism.exactBytesIdentical -ne $true) {
            throw "Committed $platform OfficeIMO evidence is not byte deterministic."
        }

        $outputs = @($engine.outputs)
        if ($outputs.Count -ne 3) {
            throw "Committed $platform $engineName evidence must contain three fresh-worker outputs."
        }
        foreach ($output in $outputs) {
            if ($output.durationMilliseconds -le 0 -or
                $output.sizeBytes -le 0 -or
                $output.managedAllocatedBytes -le 0 -or
                [string] $output.sha256 -notmatch '^[0-9a-f]{64}$' -or
                [string] $output.semanticSha256 -notmatch '^[0-9a-f]{64}$' -or
                $output.contract.pageCount -ne $report.input.expectedPageCount -or
                $output.contract.reportMarkerCount -ne $report.input.expectedReportMarkerCount -or
                $output.contract.textLength -le 0 -or
                $output.contract.tagged -ne $true -or
                $output.contract.marked -ne $true -or
                $output.contract.structureElementCount -le 0 -or
                $output.contract.markedContentReferenceCount -le 0 -or
                $output.contract.parentTreeEntryCount -ne $report.input.expectedPageCount -or
                $output.contract.hasDocumentStructureElement -ne $true -or
                $output.contract.figuresHaveAlternateText -ne $true -or
                $output.managedVisual.sizeBytes -le 0 -or
                $output.managedVisual.sha256 -notmatch '^[0-9a-f]{64}$' -or
                $output.externalVisual.sizeBytes -le 0 -or
                $output.externalVisual.sha256 -notmatch '^[0-9a-f]{64}$' -or
                $output.externalVisual.renderer -notmatch '^pdftoppm version ' -or
                @($output.externalVisual.diagnostics).Count -ne 0 -or
                $output.processTreeMemory.peakWorkingSetBytes -le 0 -or
                $output.processTreeMemory.sampleCount -le 0 -or
                $output.processTreeMemory.minimumObservedProcessCount -le 0 -or
                $output.processTreeMemory.maximumObservedProcessCount -lt $output.processTreeMemory.minimumObservedProcessCount) {
                throw "Committed $platform $engineName iteration $($output.iteration) failed the artifact contract."
            }

            if ($null -eq $commonTextLength) {
                $commonTextLength = [int] $output.contract.textLength
            } elseif ($commonTextLength -ne [int] $output.contract.textLength) {
                throw 'Equivalent-work HTML/PDF evidence did not retain the same normalized text length across engines and platforms.'
            }
        }
    }
}

& git -C $repositoryRoot cat-file -e "$officeCommit`^{commit}"
if ($LASTEXITCODE -ne 0) { throw "Measured OfficeIMO artifact-evidence commit $officeCommit is unavailable." }
& git -C $repositoryRoot merge-base --is-ancestor $officeCommit HEAD
$measuredCommitIsAncestor = $LASTEXITCODE -eq 0

$measuredPaths = @(
    'OfficeIMO.Core',
    'OfficeIMO.Html',
    'OfficeIMO.Html.Pdf',
    ':(glob)OfficeIMO.Html.Pdf.Browser/**/*.cs',
    'OfficeIMO.Pdf',
    'OfficeIMO.Pdf.Benchmarks.Comparisons'
)
& git -C $repositoryRoot diff --quiet $officeCommit HEAD -- @measuredPaths
if ($LASTEXITCODE -ne 0) {
    $relationship = $measuredCommitIsAncestor ? 'after the recorded source commit' : 'relative to the non-ancestor recorded source tree'
    throw "Committed HTML/PDF artifact evidence is stale: measured production or evidence-runner sources changed $relationship."
}
if (-not $measuredCommitIsAncestor) {
    Write-Verbose "Measured OfficeIMO artifact-evidence commit $officeCommit is squash-equivalent to HEAD for every measured path."
}
& git -C $repositoryRoot diff --quiet -- @measuredPaths
if ($LASTEXITCODE -ne 0) { throw 'Measured HTML/PDF artifact sources have uncommitted changes.' }
& git -C $repositoryRoot diff --cached --quiet -- @measuredPaths
if ($LASTEXITCODE -ne 0) { throw 'Measured HTML/PDF artifact sources have staged changes.' }
$untrackedMeasuredPaths = @(& git -C $repositoryRoot ls-files --others --exclude-standard -- @measuredPaths)
if ($LASTEXITCODE -ne 0) { throw 'Unable to inspect untracked HTML/PDF artifact source files.' }
if ($untrackedMeasuredPaths.Count -ne 0) {
    throw "Measured HTML/PDF artifact sources contain untracked files: $($untrackedMeasuredPaths -join ', ')"
}

Write-Host "Current-source Windows/Linux HTML/PDF artifact evidence verified at OfficeIMO $officeCommit and HtmlTinkerX $browserCommit."
