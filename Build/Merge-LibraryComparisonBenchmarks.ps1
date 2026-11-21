#requires -Version 7.4
<#
.SYNOPSIS
Merges portable benchmark evidence bundles into the website's evidence catalog.
.DESCRIPTION
PowerForge verifies every normalized result against its recorded hash, re-evaluates every lane against
Build/benchmark-declarations.json and the standard publish policy, and keeps earlier runs of a lane in the
superseded history. The destination is also an input, so merging one platform's bundle keeps the other lanes.
With no sources and no destination it writes a catalog of declarations only, so every comparison is listed as
not measured yet. Catalogs from earlier schemas are rejected; record fresh runs instead.
.PARAMETER SourcePath
Zero or more bundle index.json files, each with its normalized results beside it.
.PARAMETER Path
Destination catalog. Defaults to the website's evidence catalog.
.PARAMETER PowerForgeRoot
Optional PowerForge source checkout at the pinned commit, used by the shared initializer.
.PARAMETER DryRun
Returns the merge plan without setup, builds or filesystem writes.
.EXAMPLE
./Build/Merge-LibraryComparisonBenchmarks.ps1 -SourcePath ./artifacts/benchmark-bundles/windows/index.json, ./artifacts/benchmark-bundles/macos/index.json
.OUTPUTS
A merge plan in DryRun mode; otherwise the merged PowerForge catalog.
#>
[CmdletBinding()]
param(
    [string[]] $SourcePath = @(),
    [string] $Path = (Join-Path $PSScriptRoot '../Website/static/data/benchmarks/evidence/index.json'),
    [string] $PowerForgeRoot = $env:POWERFORGE_ROOT,
    [switch] $DryRun
)

$ErrorActionPreference = 'Stop'
$resultPathPrefix = '/data/benchmarks/evidence'
$declarationPath = Join-Path $PSScriptRoot 'benchmark-declarations.json'
$Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
& (Join-Path $PSScriptRoot 'Export-BenchmarkDeclarations.ps1') -Check
$sources = @(
    foreach ($source in $SourcePath) {
        $resolved = (Resolve-Path -LiteralPath $source -ErrorAction Stop).ProviderPath
        if (-not (Test-Path -LiteralPath $resolved -PathType Leaf)) {
            throw "Source must be a bundle index file: '$resolved'."
        }
        $resolved
    }
    if (Test-Path -LiteralPath $Path -PathType Leaf) {
        $Path
    }
)
$sources = @($sources | Sort-Object -Unique)

if ($DryRun) {
    [pscustomobject]@{
        DryRun = $true
        SourcePath = @($sources)
        Path = $Path
        DeclarationPath = $declarationPath
        ResultPathPrefix = $resultPathPrefix
        DeclaredOnly = @($sources).Count -eq 0
        PowerForge = & (Join-Path $PSScriptRoot 'Initialize-BenchmarkPowerForge.ps1') -PowerForgeRoot $PowerForgeRoot -PlanOnly
    }
    return
}

& (Join-Path $PSScriptRoot 'Initialize-BenchmarkPowerForge.ps1') -PowerForgeRoot $PowerForgeRoot | Out-Null
. (Join-Path $PSScriptRoot 'BenchmarkEvidence.ps1')
$temporaryDirectory = Join-Path ([System.IO.Path]::GetTempPath()) ('officeimo-benchmark-merge-' + [guid]::NewGuid().ToString('N'))
try {
    if (@($sources).Count -eq 0) {
        New-Item -ItemType Directory -Path $temporaryDirectory | Out-Null
        $emptySource = Join-Path $temporaryDirectory 'index.json'
        Write-BenchmarkEvidenceResult -Path $emptySource -InputObject ([PowerForge.BenchmarkEvidenceCatalog]::new())
        $sources = @($emptySource)
    }
    New-Item -ItemType Directory -Path (Split-Path -Parent $Path) -Force | Out-Null
    $merged = Merge-BenchmarkEvidenceCatalog -SourcePath @($sources) -Path $Path -DeclarationPath $declarationPath `
        -ResultPathPrefix $resultPathPrefix -Policy ([PowerForge.BenchmarkPublishPolicy]::CreateDefault())
    # The catalog folder holds only the results of its current lanes. Files of replaced runs stay in Git history and
    # in the bundles that recorded them; the superseded history keeps their dates and hashes.
    $catalogDirectory = Split-Path -Parent $Path
    $current = @($merged.Entries | ForEach-Object { [System.IO.Path]::GetFileName([string] $_.ArtifactRelativePath) })
    Get-ChildItem -LiteralPath $catalogDirectory -File -Filter '*.json' |
        Where-Object { $_.Name -match '\.[0-9a-f]{64}\.json$' -and $_.Name -notin $current } |
        Remove-Item -ErrorAction Stop
    $merged
} finally {
    if (Test-Path -LiteralPath $temporaryDirectory) {
        Remove-Item -LiteralPath $temporaryDirectory -Recurse -ErrorAction Stop
    }
}
