#requires -Version 7.4
<#
.SYNOPSIS
Regenerates the benchmark overview from the current library comparison catalog.
.DESCRIPTION
Keeps the website pipeline entry point on the same catalog-based generator used
by the bundle merge. Writes both generated comparison partials; the legacy Excel
matrix remains separate.
.EXAMPLE
./Build/Generate-BenchmarkWebsiteOverview.ps1
.OUTPUTS
Writes benchmarks-overview.html and benchmarks-catalog.html through the generator.
#>
& (Join-Path $PSScriptRoot 'Generate-LibraryComparisonWebsiteData.ps1')
