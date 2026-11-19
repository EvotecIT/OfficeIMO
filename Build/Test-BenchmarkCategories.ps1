<#
.SYNOPSIS
Checks that the library: and case: BenchmarkDotNet categories in benchmark code match the declarations.
.DESCRIPTION
The evidence importer takes a row's library and case from its [BenchmarkCategory("library:<id>", "case:<name>")]
categories before the runner's method maps, so a category that disagrees with Build/benchmark-declarations.json would
publish a row under another identity. For every declared workload, this reads the benchmark classes its filter selects
and requires each tagged method to name the library and case the declaration derives for it. It also requires the
method its definition expects to exist. A method may give only its library (as restricted libraries do for the
benchmark library filter); its case then comes from the runner's method map.
Nothing is built or run.
.EXAMPLE
./Build/Test-BenchmarkCategories.ps1
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'BenchmarkDeclarations.ps1')

$repositoryRoot = Split-Path -Parent $PSScriptRoot
$libraries = @(Get-BenchmarkDeclaredLibraries -Policy (Get-BenchmarkLibraryPolicy))
$sourceCache = @{}

function Get-BenchmarkSourceMethods {
    param([string] $ProjectDirectory)

    if ($sourceCache.ContainsKey($ProjectDirectory)) { return $sourceCache[$ProjectDirectory] }
    $methods = [System.Collections.Generic.List[object]]::new()
    foreach ($file in Get-ChildItem -LiteralPath $ProjectDirectory -Filter '*.cs' -File -Recurse |
            Where-Object { $_.FullName -notmatch '[\\/](bin|obj)[\\/]' }) {
        $text = Get-Content -LiteralPath $file.FullName -Raw
        $namespace = if ($text -match '(?m)^\s*namespace\s+([\w\.]+)') { $Matches[1] } else { '' }
        $class = $null
        $classCategories = @()
        $pending = [System.Collections.Generic.List[string]]::new()
        foreach ($line in $text -split "`r?`n") {
            $trimmed = $line.Trim()
            if ($trimmed.StartsWith('[')) { $pending.Add($trimmed); continue }
            if ($trimmed.StartsWith('//') -or $trimmed.Length -eq 0) { continue }
            if ($trimmed -match '\b(?:class|record)\s+(\w+)') {
                $class = $Matches[1]
                $classCategories = @($pending | ForEach-Object { [regex]::Matches($_, 'BenchmarkCategory\(([^)]*)\)') } | ForEach-Object {
                        [regex]::Matches($_.Groups[1].Value, '"([^"]*)"') | ForEach-Object { $_.Groups[1].Value } })
                $pending.Clear()
                continue
            }
            if ($null -ne $class -and @($pending | Where-Object { $_ -match '^\[Benchmark\b(?!Category)' }).Count -gt 0 -and
                $trimmed -match '\b(\w+)\s*\(') {
                $categories = @($pending | ForEach-Object { [regex]::Matches($_, 'BenchmarkCategory\(([^)]*)\)') } | ForEach-Object {
                        [regex]::Matches($_.Groups[1].Value, '"([^"]*)"') | ForEach-Object { $_.Groups[1].Value } })
                $methods.Add([pscustomobject]@{
                        FullName = (@($namespace, $class, $Matches[1]) | Where-Object { $_ }) -join '.'
                        Method = $Matches[1]
                        Categories = @($classCategories + $categories)
                        File = $file.FullName
                    })
            }
            $pending.Clear()
        }
    }
    $sourceCache[$ProjectDirectory] = $methods
    $methods
}

function Get-CategoryValue {
    param([string[]] $Categories, [string] $Prefix)
    $values = @($Categories | Where-Object { $_.StartsWith($Prefix, [StringComparison]::OrdinalIgnoreCase) } |
        ForEach-Object { $_.Substring($Prefix.Length).Trim() } | Sort-Object -Unique)
    if ($values.Count -gt 1) { throw "Categories name more than one value for '$Prefix': $($values -join ', ')." }
    if ($values.Count -eq 1) { $values[0] } else { $null }
}

$problems = [System.Collections.Generic.List[string]]::new()
$definitions = Get-LibraryComparisonDefinitions -Framework net10.0
foreach ($workload in $definitions.Keys) {
    $definition = $definitions[$workload]
    $projectDirectory = Join-Path $repositoryRoot (Split-Path -Parent ($definition.Project -replace '\\', '/'))
    $pattern = '^' + ([regex]::Escape(($definition.Filter -replace '\(.*$', '')) -replace '\\\*', '.*') + '$'
    $selected = @(Get-BenchmarkSourceMethods -ProjectDirectory $projectDirectory | Where-Object { $_.FullName -match $pattern })
    $expectedMethods = @($definition.ExpectedCases | ForEach-Object { (ConvertFrom-BenchmarkExpectedCase $_).Method } | Sort-Object -Unique)
    foreach ($method in $expectedMethods) {
        if (@($selected | Where-Object Method -ceq $method).Count -eq 0) {
            $problems.Add("${workload}: the filter '$($definition.Filter)' selects no benchmark method '$method' in $projectDirectory.")
        }
    }
    $tagged = @($selected | Where-Object { $null -ne (Get-CategoryValue $_.Categories 'library:') -or $null -ne (Get-CategoryValue $_.Categories 'case:') })
    foreach ($item in $tagged) {
        if ($item.Method -cnotin $expectedMethods) { continue }
        $identity = Get-BenchmarkMethodIdentity -Workload $workload -Definition $definition -Libraries $libraries -Method $item.Method
        $library = Get-CategoryValue $item.Categories 'library:'
        $case = Get-CategoryValue $item.Categories 'case:'
        if (($null -ne $library -and $library -cne $identity.Library) -or ($null -ne $case -and $case -cne $identity.Case)) {
            $problems.Add("${workload}: $($item.Method) is tagged library:$library, case:$case; the declaration derives library:$($identity.Library), case:$($identity.Case).")
        }
    }
}

if ($problems.Count -gt 0) {
    throw ("Benchmark categories disagree with the declarations:" + [Environment]::NewLine + ($problems -join [Environment]::NewLine))
}
[pscustomobject]@{
    Workloads = $definitions.Count
    Tagged = @($definitions.Keys | Where-Object {
            $definition = $definitions[$_]
            $projectDirectory = Join-Path $repositoryRoot (Split-Path -Parent ($definition.Project -replace '\\', '/'))
            @(Get-BenchmarkSourceMethods -ProjectDirectory $projectDirectory | Where-Object {
                    $null -ne (Get-CategoryValue $_.Categories 'library:') -and
                    $_.FullName -match ('^' + ([regex]::Escape(($definition.Filter -replace '\(.*$', '')) -replace '\\\*', '.*') + '$')
                }).Count -gt 0
        }).Count
}
