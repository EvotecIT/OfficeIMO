#requires -Version 7.4
<#
.SYNOPSIS
Validates and copies the published comparison catalog for build-time templates.
.DESCRIPTION
The catalog is the only evidence input. This does not measure workloads, promote
quick results, fill missing environments, or update the legacy Excel matrix.
.PARAMETER CatalogPath
Merged PowerForge v4 index.json to validate and copy.
.PARAMETER WebsiteIndexPath
Build-time template data file receiving the exact published catalog bytes.
.PARAMETER DryRun
Validates the catalog and returns output paths without writing files.
.EXAMPLE
./Build/Generate-LibraryComparisonWebsiteData.ps1 -CatalogPath ./Website/static/data/benchmarks/library-comparisons/index.json
.OUTPUTS
An output plan in DryRun mode. Otherwise writes an identical catalog copy.
#>
[CmdletBinding()]
param(
    [string] $CatalogPath = (Join-Path $PSScriptRoot '../Website/static/data/benchmarks/library-comparisons/index.json'),
    [string] $WebsiteIndexPath = (Join-Path $PSScriptRoot '../Website/data/benchmarks_index.json'),
    [switch] $DryRun
)

$ErrorActionPreference = 'Stop'
$CatalogPath = (Resolve-Path -LiteralPath $CatalogPath -ErrorAction Stop).ProviderPath
$WebsiteIndexPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($WebsiteIndexPath)
$catalogBytes = [System.IO.File]::ReadAllBytes($CatalogPath)
$json = [System.Text.Encoding]::UTF8.GetString($catalogBytes).TrimStart([char] 0xFEFF)
if (-not (Test-Json -Json $json -SchemaFile (Join-Path $PSScriptRoot 'library-comparisons.schema.json'))) {
    throw "Invalid v4 comparison catalog: '$CatalogPath'."
}
$catalog = $json | ConvertFrom-Json -Depth 100
if ($DryRun) {
    [pscustomobject]@{
        DryRun = $true
        CatalogPath = $CatalogPath
        OutputPath = @($WebsiteIndexPath)
        ComparisonCount = @($catalog.comparisons).Count
        EntryCount = @($catalog.entries).Count
    }
    return
}

New-Item -ItemType Directory -Path (Split-Path -Parent $WebsiteIndexPath) -Force | Out-Null
# Retain the source timestamp strings and numeric representation on every locale.
[System.IO.File]::WriteAllBytes($WebsiteIndexPath, $catalogBytes)
Write-Host "Copied comparison catalog to '$WebsiteIndexPath'."
