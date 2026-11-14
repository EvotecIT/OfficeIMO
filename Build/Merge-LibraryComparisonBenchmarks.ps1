#requires -Version 7.4
<#
.SYNOPSIS
Merges portable library comparison bundles and refreshes build-time website data.
.DESCRIPTION
PowerForge verifies and merges the normalized artifacts. The destination is also
an input so collecting a subset does not discard other platform or workload lanes.
Declarations come from both framework plans, including restricted workloads.
.PARAMETER SourcePath
Zero or more bundle index.json files, each accompanied by its result artifacts.
.PARAMETER Path
Destination index. Defaults to the website's library comparison catalog.
.PARAMETER PowerForgeRoot
Optional exact PowerForge source checkout used by the shared initializer.
.PARAMETER WebsiteIndexPath
Build-time template data file receiving an identical copy of the published index.
.PARAMETER DryRun
Returns the merge plan without setup, builds, measurements or filesystem writes.
.EXAMPLE
./Build/Merge-LibraryComparisonBenchmarks.ps1 -SourcePath ./artifacts/benchmark-bundles/windows/index.json, ./artifacts/benchmark-bundles/linux/index.json, ./artifacts/benchmark-bundles/macos/index.json
.OUTPUTS
A merge plan in DryRun mode; otherwise the merged PowerForge catalog.
#>
[CmdletBinding()]
param(
    [string[]] $SourcePath = @(),
    [string] $Path = (Join-Path $PSScriptRoot '../Website/static/data/benchmarks/library-comparisons/index.json'),
    [string] $PowerForgeRoot = $env:POWERFORGE_ROOT,
    [string] $WebsiteIndexPath = (Join-Path $PSScriptRoot '../Website/data/benchmarks_index.json'),
    [switch] $DryRun
)

$ErrorActionPreference = 'Stop'
$resultPathPrefix = '/data/benchmarks/library-comparisons'
$Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
$WebsiteIndexPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($WebsiteIndexPath)
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

. (Join-Path $PSScriptRoot 'LibraryComparisonDefinitions.ps1')
$declarations = [ordered]@{}
foreach ($framework in @('net8.0', 'net10.0')) {
    $plan = & (Join-Path $PSScriptRoot 'Run-LibraryComparisonBenchmarks.ps1') `
        -Workload all -RunMode full -Framework $framework -PlanOnly
    $frameworkDefinitions = Get-LibraryComparisonDefinitions -Framework $framework
    foreach ($item in $plan) {
        $supported = $frameworkDefinitions[$item.Workload].Frameworks
        if ($supported -and $framework -notin $supported) { continue }
        $declarations[$item.ComparisonId] = [pscustomobject]@{
            comparisonId = $item.ComparisonId
            area = $item.Area
            title = $item.Title
            libraries = @($item.Libraries)
            expectedPlatforms = @($item.ExpectedPlatforms)
        }
    }
}

# Historical placement and warmed profiles remain distinct comparisons. Map only
# the explicit suffixes produced by the runner; never infer an area from a name.
foreach ($source in $sources) {
    $catalog = Get-Content -LiteralPath $source -Raw -Encoding UTF8 | ConvertFrom-Json -Depth 100
    foreach ($entry in @($catalog.entries)) {
        $id = [string] $entry.comparisonId
        if ($declarations.Contains($id)) {
            continue
        }
        $baseId = $id -replace '-affinity-0x[0-9a-fA-F]+$', '' -replace '-warmed-v1$', ''
        if ($baseId -eq $id -or -not $declarations.Contains($baseId)) {
            throw "Comparison '$id' has no runner declaration. Add its definition before merging."
        }
        $definition = $declarations[$baseId]
        $declarations[$id] = [pscustomobject]@{
            comparisonId = $id
            area = $definition.area
            title = $definition.title + ' (' + $id.Substring($baseId.Length).TrimStart('-') + ')'
            libraries = @($definition.libraries)
            expectedPlatforms = @($definition.expectedPlatforms)
        }
    }
}
$definitions = @($declarations.Values | Sort-Object comparisonId)

if ($DryRun) {
    $powerForge = & (Join-Path $PSScriptRoot 'Initialize-BenchmarkPowerForge.ps1') -PowerForgeRoot $PowerForgeRoot -PlanOnly
    [pscustomobject]@{
        DryRun = $true
        SourcePath = @($sources)
        Path = $Path
        ResultPathPrefix = $resultPathPrefix
        DeclaredOnly = @($sources).Count -eq 0
        Comparisons = $definitions
        WebsiteIndexPath = $WebsiteIndexPath
        GeneratedFiles = @('benchmarks_index.json')
        PowerForge = $powerForge
    }
    return
}

& (Join-Path $PSScriptRoot 'Initialize-BenchmarkPowerForge.ps1') -PowerForgeRoot $PowerForgeRoot | Out-Null
# Declaration inputs are transient merge metadata, separate from raw provenance.
$temporaryDirectory = Join-Path ([System.IO.Path]::GetTempPath()) ('officeimo-comparison-merge-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $temporaryDirectory | Out-Null
try {
    $comparisonPath = Join-Path $temporaryDirectory 'comparisons.json'
    ConvertTo-Json -InputObject $definitions -Depth 20 |
        Set-Content -LiteralPath $comparisonPath -Encoding utf8NoBOM
    if (@($sources).Count -eq 0) {
        $emptySource = Join-Path $temporaryDirectory 'index.json'
        [pscustomobject]@{
            schemaVersion = 4
            expectedPlatforms = @('windows', 'linux', 'macos')
            areas = @()
            comparisons = @()
            entries = @()
            availability = @()
        } | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $emptySource -Encoding utf8NoBOM
        $sources = @($emptySource)
    }
    $merged = Merge-BenchmarkEvidenceCatalog -SourcePath @($sources) -Path $Path -ComparisonPath $comparisonPath `
        -ResultPathPrefix $resultPathPrefix
    & (Join-Path $PSScriptRoot 'Generate-LibraryComparisonWebsiteData.ps1') `
        -CatalogPath $Path `
        -WebsiteIndexPath $WebsiteIndexPath
    $merged
} finally {
    Remove-Item -LiteralPath $temporaryDirectory -Recurse -ErrorAction Stop
}
