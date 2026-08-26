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
$actualOsFamily = if ([string] $report.environment.osDescription -match '(?i)windows') { 'Windows' } else { 'Linux' }
if ($actualOsFamily -ne $expectedOsFamily -or
    [string]::IsNullOrWhiteSpace([string] $report.environment.externalRasterizer)) {
    throw "HTML/PDF artifact evidence is not a $expectedOsFamily run with an external rasterizer."
}

$artifacts = [System.Collections.Generic.List[object]]::new()
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

foreach ($engine in @($report.engines)) {
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
