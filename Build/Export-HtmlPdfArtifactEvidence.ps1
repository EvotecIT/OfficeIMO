param(
    [Parameter(Mandatory)]
    [string] $EvidencePath,

    [Parameter(Mandatory)]
    [ValidateSet('windows', 'linux')]
    [string] $Platform,

    [Parameter(Mandatory)]
    [string] $OutputPath
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$resolvedEvidencePath = (Resolve-Path -LiteralPath $EvidencePath).Path
$reportPath = if (Test-Path -LiteralPath $resolvedEvidencePath -PathType Container) {
    Join-Path $resolvedEvidencePath 'html-pdf-evidence.json'
} else {
    $resolvedEvidencePath
}
$evidenceRoot = Split-Path -Parent $reportPath
$pathComparison = if ([System.Environment]::OSVersion.Platform -eq [System.PlatformID]::Win32NT) {
    [System.StringComparison]::OrdinalIgnoreCase
} else {
    [System.StringComparison]::Ordinal
}
$report = Get-Content -LiteralPath $reportPath -Raw | ConvertFrom-Json

if ($report.schemaVersion -ne 2 -or
    [string] $report.scale -ne 'High' -or
    [int] $report.iterations -lt 3) {
    throw 'HTML/PDF artifact evidence must be a schema-v2 High-scale run with at least three iterations.'
}

$expectedOsFamily = $Platform -eq 'windows' ? 'Windows' : 'Linux'
$osDescription = [string] $report.environment.osDescription
$actualOsFamily = if ($osDescription -match '(?i)windows') {
    'Windows'
} elseif ($osDescription -match '(?i)linux') {
    'Linux'
} else {
    'Unknown'
}
if ($actualOsFamily -ne $expectedOsFamily -or
    [string]::IsNullOrWhiteSpace([string] $report.environment.externalRasterizer)) {
    throw "HTML/PDF artifact evidence is not a $expectedOsFamily run with an external rasterizer."
}

$expectedEngines = @('Chromium', 'ITextPdfHtml', 'OfficeIMO', 'PeachPDF')
$engines = @($report.engines)
$actualEngines = @($engines | ForEach-Object { [string] $_.engine } | Sort-Object)
if ($engines.Count -ne $expectedEngines.Count -or
    @(Compare-Object -ReferenceObject $expectedEngines -DifferenceObject $actualEngines).Count -ne 0) {
    throw 'HTML/PDF artifact evidence must contain exactly the required Chromium, ITextPdfHtml, OfficeIMO, and PeachPDF engines.'
}

foreach ($engine in $engines) {
    $engineName = [string] $engine.engine
    $outputs = @($engine.outputs)
    $iterations = @($outputs | ForEach-Object { [int] $_.iteration } | Sort-Object -Unique)
    $expectedIterations = @(1..([int] $report.iterations))
    if ($outputs.Count -ne [int] $report.iterations -or
        $iterations.Count -ne $expectedIterations.Count -or
        @(Compare-Object -ReferenceObject $expectedIterations -DifferenceObject $iterations).Count -ne 0) {
        throw "HTML/PDF artifact evidence engine '$engineName' must contain exactly one output for every declared iteration."
    }

    if ([string] $engine.cancellation.status -notin @('Passed', 'Unsupported') -or
        $engine.memoryComparable -ne $true -or
        $engine.determinism.exactBytesIdentical -ne $true -or
        $engine.determinism.semanticOutputIdentical -ne $true -or
        $engine.determinism.managedVisualPreviewIdentical -ne $true -or
        $engine.determinism.externalVisualPreviewIdentical -ne $true) {
        throw "HTML/PDF artifact evidence engine '$engineName' does not satisfy the cancellation, memory, and determinism contract."
    }

    foreach ($output in $outputs) {
        $contract = $output.contract
        if ($null -eq $contract -or
            [int] $contract.pageCount -lt 1 -or
            [int] $contract.textLength -lt 1 -or
            [int] $contract.reportMarkerCount -lt 1 -or
            $contract.tagged -ne $true -or
            $contract.marked -ne $true -or
            [string]::IsNullOrWhiteSpace([string] $contract.catalogLanguage) -or
            [int] $contract.structureElementCount -lt 1 -or
            [int] $contract.markedContentReferenceCount -lt 1 -or
            [int] $contract.parentTreeEntryCount -lt 1 -or
            $contract.hasDocumentStructureElement -ne $true -or
            $contract.figuresHaveAlternateText -ne $true -or
            [string] $output.semanticSha256 -notmatch '^[0-9a-f]{64}$') {
            throw "HTML/PDF artifact evidence engine '$engineName' contains an output without the required semantic and tagged-PDF contract."
        }
    }
}

$artifacts = [System.Collections.Generic.List[object]]::new()
function Assert-NoArtifactPathLinks {
    param(
        [Parameter(Mandatory)][string] $RootPath,
        [Parameter(Mandatory)][string] $FullPath,
        [Parameter(Mandatory)][string] $RelativePath
    )

    $rootPathFull = [System.IO.Path]::GetFullPath($RootPath).TrimEnd(
        [System.IO.Path]::DirectorySeparatorChar,
        [System.IO.Path]::AltDirectorySeparatorChar)
    $relativeFromRoot = [System.IO.Path]::GetRelativePath($rootPathFull, $FullPath)
    $currentPath = $rootPathFull
    $pathsToInspect = [System.Collections.Generic.List[string]]::new()
    $pathsToInspect.Add($currentPath)
    foreach ($segment in $relativeFromRoot.Split(
            [char[]]@(
                [System.IO.Path]::DirectorySeparatorChar,
                [System.IO.Path]::AltDirectorySeparatorChar),
            [System.StringSplitOptions]::RemoveEmptyEntries)) {
        $currentPath = Join-Path $currentPath $segment
        $pathsToInspect.Add($currentPath)
    }

    foreach ($path in $pathsToInspect) {
        $item = Get-Item -LiteralPath $path -Force
        $isReparsePoint = ($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0
        $isLink = $item.PSObject.Properties.Name -contains 'LinkType' -and
            -not [string]::IsNullOrWhiteSpace([string] $item.LinkType)
        if ($isReparsePoint -or $isLink) {
            throw "Artifact path contains a symbolic link or reparse point: $RelativePath"
        }
    }
}

function Add-ValidatedArtifact {
    param(
        [Parameter(Mandatory)][string] $Kind,
        [Parameter(Mandatory)][string] $RelativePath,
        [Parameter(Mandatory)][long] $ExpectedSize,
        [Parameter(Mandatory)][string] $ExpectedSha256
    )

    $fullPath = [System.IO.Path]::GetFullPath((Join-Path $evidenceRoot $RelativePath))
    $rootPrefix = [System.IO.Path]::GetFullPath($evidenceRoot).TrimEnd(
        [System.IO.Path]::DirectorySeparatorChar,
        [System.IO.Path]::AltDirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar
    if (-not $fullPath.StartsWith($rootPrefix, $pathComparison)) {
        throw "Artifact path escapes the evidence root: $RelativePath"
    }
    if (-not (Test-Path -LiteralPath $fullPath -PathType Leaf)) {
        throw "Artifact is missing: $RelativePath"
    }
    Assert-NoArtifactPathLinks `
        -RootPath $evidenceRoot `
        -FullPath $fullPath `
        -RelativePath $RelativePath

    $item = Get-Item -LiteralPath $fullPath
    $actualHash = (Get-FileHash -LiteralPath $fullPath -Algorithm SHA256).Hash.ToLowerInvariant()
    if ($item.Length -ne $ExpectedSize -or
        -not [string]::Equals($actualHash, $ExpectedSha256, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Artifact size or SHA-256 does not match the evidence report: $RelativePath"
    }

    $artifacts.Add([ordered]@{
            kind = $Kind
            relativePath = $RelativePath.Replace('\', '/')
            sizeBytes = $item.Length
            sha256 = $actualHash
        })
}

Add-ValidatedArtifact `
    -Kind 'input' `
    -RelativePath ([string] $report.input.relativePath) `
    -ExpectedSize ([long] $report.input.sizeBytes) `
    -ExpectedSha256 ([string] $report.input.sha256)

foreach ($engine in $engines) {
    foreach ($output in @($engine.outputs)) {
        Add-ValidatedArtifact `
            -Kind ("pdf:$($engine.engine)") `
            -RelativePath ([string] $output.relativePath) `
            -ExpectedSize ([long] $output.sizeBytes) `
            -ExpectedSha256 ([string] $output.sha256)
        Add-ValidatedArtifact `
            -Kind ("managed-preview:$($engine.engine)") `
            -RelativePath ([string] $output.managedVisual.relativePath) `
            -ExpectedSize ([long] $output.managedVisual.sizeBytes) `
            -ExpectedSha256 ([string] $output.managedVisual.sha256)
        if ($null -eq $output.externalVisual) {
            throw "External visual evidence is missing for $($engine.engine) iteration $($output.iteration)."
        }
        Add-ValidatedArtifact `
            -Kind ("external-preview:$($engine.engine)") `
            -RelativePath ([string] $output.externalVisual.relativePath) `
            -ExpectedSize ([long] $output.externalVisual.sizeBytes) `
            -ExpectedSha256 ([string] $output.externalVisual.sha256)
    }
}

$orderedArtifacts = @($artifacts | Sort-Object { $_.relativePath })
if (@($orderedArtifacts | Group-Object { $_.relativePath } | Where-Object { $_.Count -ne 1 }).Count -ne 0) {
    throw 'HTML/PDF artifact evidence contains duplicate artifact paths.'
}
$manifestText = ($orderedArtifacts | ForEach-Object {
        "$($_.kind)|$($_.relativePath)|$($_.sizeBytes)|$($_.sha256)"
    }) -join "`n"
$manifestBytes = [System.Text.Encoding]::UTF8.GetBytes($manifestText)
$manifestHash = [Convert]::ToHexString(
    [System.Security.Cryptography.SHA256]::HashData($manifestBytes)).ToLowerInvariant()

$summary = [ordered]@{
    schemaVersion = 1
    format = 'officeimo.html-pdf-artifact-evidence-summary'
    platform = $Platform
    artifactBundle = [ordered]@{
        artifactCount = $orderedArtifacts.Count
        totalBytes = [long] (@($orderedArtifacts | ForEach-Object { $_.sizeBytes } | Measure-Object -Sum).Sum)
        manifestSha256 = $manifestHash
        artifacts = $orderedArtifacts
    }
    report = $report
}

$resolvedOutputPath = [System.IO.Path]::GetFullPath($OutputPath)
New-Item -ItemType Directory -Path (Split-Path -Parent $resolvedOutputPath) -Force | Out-Null
$json = ($summary | ConvertTo-Json -Depth 100).Replace("`r`n", "`n") + "`n"
[System.IO.File]::WriteAllText($resolvedOutputPath, $json, [System.Text.UTF8Encoding]::new($false))
Write-Host "Validated HTML/PDF artifact evidence summary written to '$resolvedOutputPath'."
