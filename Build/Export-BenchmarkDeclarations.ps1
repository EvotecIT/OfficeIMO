<#
.SYNOPSIS
Writes the benchmark declarations that the evidence catalog and the runner check results against.
.DESCRIPTION
Generates Build/benchmark-declarations.json (PowerForge schema 5) from the runner definitions in
Build/LibraryComparisonDefinitions.ps1 and the library policy in Build/benchmark-libraries.json. Never edit the
generated file; change the definitions or the policy and run this script. -Check fails when the committed file is
out of date, without writing it.
.PARAMETER Path
Declaration file to write or check.
.PARAMETER Check
Compares the generated document with the file instead of writing it.
.EXAMPLE
./Build/Export-BenchmarkDeclarations.ps1
.EXAMPLE
./Build/Export-BenchmarkDeclarations.ps1 -Check
#>
param(
    [string] $Path = (Join-Path $PSScriptRoot 'benchmark-declarations.json'),
    [switch] $Check
)

$ErrorActionPreference = 'Stop'
# .NET file APIs resolve relative paths against the process directory, not the PowerShell location.
$Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
. (Join-Path $PSScriptRoot 'BenchmarkDeclarations.ps1')

$document = New-BenchmarkDeclarationDocument
$json = ($document | ConvertTo-Json -Depth 12) -replace "`r`n", "`n"
$json += "`n"

if ($Check) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Benchmark declarations '$Path' are missing; run ./Build/Export-BenchmarkDeclarations.ps1."
    }
    $current = (Get-Content -LiteralPath $Path -Raw -Encoding UTF8) -replace "`r`n", "`n"
    if ($current -cne $json) {
        throw "Benchmark declarations '$Path' are out of date; run ./Build/Export-BenchmarkDeclarations.ps1."
    }
    return
}

[System.IO.File]::WriteAllText($Path, $json, [System.Text.UTF8Encoding]::new($false))
[pscustomobject]@{
    Path = $Path
    Libraries = @($document.libraries).Count
    Comparisons = @($document.comparisons).Count
}
