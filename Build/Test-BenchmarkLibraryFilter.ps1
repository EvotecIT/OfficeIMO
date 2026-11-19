<#
.SYNOPSIS
Checks that the restricted-library benchmark filter keeps open libraries and drops restricted ones unless enabled.
.DESCRIPTION
Benchmarks/BenchmarkLibraryFilter.cs reads the library policy embedded from Build/benchmark-libraries.json. This lists
the 65K XLSX comparison of OfficeIMO.Excel.Benchmarks, whose OfficeIMO case carries library:OfficeIMO (open) and whose
EPPlus case carries library:epplus (acceptance-required), and requires:
- without OFFICEIMO_BENCHMARK_ENABLED_LIBRARIES, the OfficeIMO case is listed and the EPPlus case is not;
- with OFFICEIMO_BENCHMARK_ENABLED_LIBRARIES=epplus, both are listed.
OfficeIMO.Excel.Benchmarks has no benchmarks on net472, so for net472 it lists the Word read comparison of
OfficeIMO.Word.Benchmarks instead: its OfficeIMO case and its DocX case (library:DocX, acceptance-required) under the
same rules, with docx as the enabled library. That covers the .NET Framework policy embed and JSON parsing.
Listing selects cases only; nothing is measured and no restricted library runs.
.PARAMETER Framework
Target framework of the benchmark build to list.
.PARAMETER NoBuild
Uses an existing Release build instead of building the project first.
.EXAMPLE
./Build/Test-BenchmarkLibraryFilter.ps1
#>
[CmdletBinding()]
param(
    [ValidateSet('net8.0', 'net10.0', 'net472')]
    [string] $Framework = 'net10.0',
    [switch] $NoBuild
)

$ErrorActionPreference = 'Stop'
$repositoryRoot = Split-Path -Parent $PSScriptRoot
$projectName, $officeImoCase, $restrictedCase, $restrictedId = if ($Framework -eq 'net472') {
    'OfficeIMO.Word.Benchmarks', 'OfficeIMO.Word.Benchmarks.WordReadComparisonBenchmarks.OfficeIMO',
    'OfficeIMO.Word.Benchmarks.WordReadComparisonBenchmarks.DocX', 'docx'
} else {
    'OfficeIMO.Excel.Benchmarks', 'OfficeIMO.Excel.Benchmarks.MarkPflug65KXlsxBenchmarks.OfficeIMO',
    'OfficeIMO.Excel.Benchmarks.MarkPflug65KXlsxBenchmarks.EPPlus', 'epplus'
}
$project = Join-Path $repositoryRoot "$projectName/$projectName.csproj"
if (-not $NoBuild) {
    & dotnet build $project -c Release -f $Framework -nologo -v quiet
    if ($LASTEXITCODE -ne 0) { throw "Building $project failed." }
}
$extension = if ($Framework -eq 'net472') { 'exe' } else { 'dll' }
$assembly = Join-Path $repositoryRoot "$projectName/bin/Release/$Framework/$projectName.$extension"
if (-not (Test-Path -LiteralPath $assembly -PathType Leaf)) { throw "No Release build at $assembly; run without -NoBuild." }

function Get-ListedCases {
    param([string] $EnabledLibraries)

    $previous = $env:OFFICEIMO_BENCHMARK_ENABLED_LIBRARIES
    try {
        $env:OFFICEIMO_BENCHMARK_ENABLED_LIBRARIES = $EnabledLibraries
        $listArguments = @('--list', 'flat', '--filter', $officeImoCase, $restrictedCase)
        $output = @(if ($extension -eq 'exe') { & $assembly @listArguments 2>&1 } else { & dotnet $assembly @listArguments 2>&1 }) |
            ForEach-Object { "$_".Trim() }
        if ($LASTEXITCODE -ne 0) { throw "Listing benchmarks failed: $($output -join [Environment]::NewLine)" }
        $output
    } finally {
        $env:OFFICEIMO_BENCHMARK_ENABLED_LIBRARIES = $previous
    }
}

$problems = [System.Collections.Generic.List[string]]::new()
$default = Get-ListedCases -EnabledLibraries $null
if ($officeImoCase -notin $default) { $problems.Add("The OfficeIMO case was dropped although OfficeIMO is open.") }
if ($restrictedCase -in $default) { $problems.Add("The restricted $restrictedId case ran although it was not enabled.") }
$enabled = Get-ListedCases -EnabledLibraries $restrictedId
if ($officeImoCase -notin $enabled) { $problems.Add("The OfficeIMO case was dropped with $restrictedId enabled.") }
if ($restrictedCase -notin $enabled) { $problems.Add("The restricted $restrictedId case was dropped although it was enabled.") }

if ($problems.Count -gt 0) {
    throw ("The restricted-library benchmark filter is wrong:" + [Environment]::NewLine + ($problems -join [Environment]::NewLine))
}
[pscustomobject]@{ Framework = $Framework; OpenCaseKept = $true; RestrictedCaseDroppedUnlessEnabled = $true }
