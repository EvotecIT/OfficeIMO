<#
.SYNOPSIS
Refreshes the CSV and Excel benchmark tables embedded in project READMEs from the benchmark evidence catalog.

.DESCRIPTION
Each marker-delimited block shows the published run of its comparisons on one platform and framework, rendered by
PowerForge's Update-BenchmarkDocument from Website/static/data/benchmarks/evidence/index.json. A comparison without
a published run on that platform shows a "no published run yet" line, so a README never quotes diagnostic or
unpublished numbers. Measure with Build/Run-LibraryComparisonBenchmarks.ps1 and merge the results first; this
script measures nothing.

.PARAMETER Platform
Platform whose published runs are shown: windows, linux or macos.

.PARAMETER Framework
Target framework whose published runs are shown.

.PARAMETER CatalogPath
Evidence catalog to render from.

.PARAMETER PowerForgeRoot
Optional PowerForge source checkout at the pinned commit, used by the shared initializer.

.EXAMPLE
./Build/Benchmarks/Update-BenchmarkReadmes.ps1

.EXAMPLE
./Build/Benchmarks/Update-BenchmarkReadmes.ps1 -Platform macos
#>
[CmdletBinding()]
param(
    [ValidateSet('windows', 'linux', 'macos')]
    [string] $Platform = 'windows',
    [ValidateSet('net8.0', 'net10.0')]
    [string] $Framework = 'net10.0',
    [string] $CatalogPath = (Join-Path $PSScriptRoot '../../Website/static/data/benchmarks/evidence/index.json'),
    [string] $PowerForgeRoot = $env:POWERFORGE_ROOT
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repositoryRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$CatalogPath = (Resolve-Path -LiteralPath $CatalogPath).ProviderPath
& (Join-Path $repositoryRoot 'Build/Initialize-BenchmarkPowerForge.ps1') -PowerForgeRoot $PowerForgeRoot | Out-Null

$blocks = @(
    @{
        Path = 'OfficeIMO.CSV.Benchmarks/README.md'
        BlockId = 'officeimo-csv-benchmark-table'
        Comparison = @('markpflug-65k-csv-decoded', 'csv-25k-datareader-write')
    }
    @{
        Path = 'OfficeIMO.Excel.Benchmarks/README.md'
        BlockId = 'officeimo-excel-benchmark-table'
        Comparison = @('markpflug-65k-xlsx-typed', 'markpflug-65k-xlsb-typed', 'markpflug-65k-xls-typed', 'xlsx-25k-datareader-write')
    }
)
foreach ($block in $blocks) {
    Update-BenchmarkDocument `
        -Path (Join-Path $repositoryRoot $block.Path) `
        -BlockId $block.BlockId `
        -Catalog $CatalogPath `
        -Comparison $block.Comparison `
        -Platform $Platform `
        -Framework $Framework
}
