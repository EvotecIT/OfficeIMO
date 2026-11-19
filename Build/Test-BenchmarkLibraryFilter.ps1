<#
.SYNOPSIS
Checks that the restricted-library benchmark filter keeps open libraries and drops restricted ones unless enabled.
.DESCRIPTION
Benchmarks/BenchmarkLibraryFilter.cs reads the library policy embedded from Build/benchmark-libraries.json. This lists
the 65K XLSX comparison of OfficeIMO.Excel.Benchmarks, whose OfficeIMO case carries library:OfficeIMO (open) and whose
EPPlus case carries library:epplus (acceptance-required), and requires:
- without OFFICEIMO_BENCHMARK_ENABLED_LIBRARIES, the OfficeIMO case is listed and the EPPlus case is not;
- with OFFICEIMO_BENCHMARK_ENABLED_LIBRARIES=epplus, both are listed.
Listing selects cases only; nothing is measured and EPPlus never runs.
.PARAMETER Framework
Target framework of the benchmark build to list.
.PARAMETER NoBuild
Uses an existing Release build instead of building the project first.
.EXAMPLE
./Build/Test-BenchmarkLibraryFilter.ps1
#>
[CmdletBinding()]
param(
    [ValidateSet('net8.0', 'net10.0')]
    [string] $Framework = 'net10.0',
    [switch] $NoBuild
)

$ErrorActionPreference = 'Stop'
$repositoryRoot = Split-Path -Parent $PSScriptRoot
$project = Join-Path $repositoryRoot 'OfficeIMO.Excel.Benchmarks/OfficeIMO.Excel.Benchmarks.csproj'
if (-not $NoBuild) {
    & dotnet build $project -c Release -f $Framework -nologo -v quiet
    if ($LASTEXITCODE -ne 0) { throw "Building $project failed." }
}
$assembly = Join-Path $repositoryRoot "OfficeIMO.Excel.Benchmarks/bin/Release/$Framework/OfficeIMO.Excel.Benchmarks.dll"
if (-not (Test-Path -LiteralPath $assembly -PathType Leaf)) { throw "No Release build at $assembly; run without -NoBuild." }

$officeImoCase = 'OfficeIMO.Excel.Benchmarks.MarkPflug65KXlsxBenchmarks.OfficeIMO'
$epplusCase = 'OfficeIMO.Excel.Benchmarks.MarkPflug65KXlsxBenchmarks.EPPlus'

function Get-ListedCases {
    param([string] $EnabledLibraries)

    $previous = $env:OFFICEIMO_BENCHMARK_ENABLED_LIBRARIES
    try {
        $env:OFFICEIMO_BENCHMARK_ENABLED_LIBRARIES = $EnabledLibraries
        $output = @(& dotnet $assembly --list flat --filter $officeImoCase $epplusCase 2>&1 | ForEach-Object { "$_".Trim() })
        if ($LASTEXITCODE -ne 0) { throw "Listing benchmarks failed: $($output -join [Environment]::NewLine)" }
        $output
    } finally {
        $env:OFFICEIMO_BENCHMARK_ENABLED_LIBRARIES = $previous
    }
}

$problems = [System.Collections.Generic.List[string]]::new()
$default = Get-ListedCases -EnabledLibraries $null
if ($officeImoCase -notin $default) { $problems.Add('The open library:OfficeIMO case was dropped although OfficeIMO is open.') }
if ($epplusCase -in $default) { $problems.Add('The restricted library:epplus case ran although EPPlus was not enabled.') }
$enabled = Get-ListedCases -EnabledLibraries 'epplus'
if ($officeImoCase -notin $enabled) { $problems.Add('The open library:OfficeIMO case was dropped with EPPlus enabled.') }
if ($epplusCase -notin $enabled) { $problems.Add('The restricted library:epplus case was dropped although EPPlus was enabled.') }

if ($problems.Count -gt 0) {
    throw ("The restricted-library benchmark filter is wrong:" + [Environment]::NewLine + ($problems -join [Environment]::NewLine))
}
[pscustomobject]@{ Framework = $Framework; OpenCaseKept = $true; RestrictedCaseDroppedUnlessEnabled = $true }
