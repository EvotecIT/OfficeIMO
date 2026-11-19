# Builds the schema 5 benchmark declarations (library policy and comparisons) from the runner definitions in
# Build/LibraryComparisonDefinitions.ps1 and the library policy in Build/benchmark-libraries.json.
# Loading this file performs no work.

. (Join-Path $PSScriptRoot 'BenchmarkLibraryPolicy.ps1')
. (Join-Path $PSScriptRoot 'LibraryComparisonDefinitions.ps1')

$script:BenchmarkDeclarationFrameworks = @('net8.0', 'net10.0')

# Methods whose name does not follow the "<Library><Case>" pattern the library and case are derived from.
# Keys are workload|method; values are library|case, with the library named as in the comparison's Libraries.
$script:BenchmarkMethodIdentityOverrides = @{
    'markdownparse|Markdig_ParseSemanticComparison_CommonMark' = 'Markdig|ParseSemantic_CommonMark'
    'ziptraverse|Platform' = 'System.IO.Compression|Default'
    'tomlloadsave|TomlynSyntax' = 'Tomlyn|Default'
    'yamlserialize|YamlDotNet' = 'YamlDotNet|Reflection'
    'yamldeserialize|YamlDotNet' = 'YamlDotNet|Reflection'
    'wordcreate|OpenXmlSdkEquivalent' = 'Open XML SDK|Default'
    'wordcreate|OpenXmlSdkMinimal' = 'Open XML SDK|Minimal'
    'wordreport|OpenXmlSdkEquivalent' = 'Open XML SDK|Default'
    'wordreport|OpenXmlSdkMinimal' = 'Open XML SDK|Minimal'
    'wordread|OneDocumentExtractText' = 'OneDocument.Documents|ExtractText'
}

function Get-BenchmarkDeclaredLibraries {
    <#
    .SYNOPSIS
    Turns the library policy into schema 5 library declarations.
    .DESCRIPTION
    Each policy library declares one package and one pinned version. A library that lists several packages or
    versions names the ones the comparisons measure in "declare"; a library built from this repository or shipped
    with the runtime is declared as such. Libraries without a pinned version (such as a later release kept only
    for local runs) are left out, so no comparison can use them.
    .OUTPUTS
    One object per declared library: Id, PolicyId, Names (policy name, id and aliases that resolve to it),
    Package, Version, Frameworks, License, LicenseUrl, Restriction and Reason.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)] [object] $Policy)

    foreach ($library in @($Policy.libraries)) {
        $declared = @(
            if ($null -ne $library.declare) {
                foreach ($item in @($library.declare)) {
                    [pscustomobject]@{ Id = [string] $item.id; Package = [string] $item.package; Version = [string] $item.version; Frameworks = @($item.frameworks | Where-Object { $_ }) }
                }
            } else {
                $packages = @($library.packages)
                $versions = @($packages | ForEach-Object { [string] $_.version } | Sort-Object -Unique)
                if ($packages.Count -eq 0) {
                    $version = switch ([string] $library.licenseSource) {
                        'repository' { 'source' }
                        'runtime' { 'runtime' }
                        default { throw "Benchmark library '$($library.id)' lists no package; declare the package it is measured from." }
                    }
                    [pscustomobject]@{ Id = [string] $library.name; Package = [string] $library.name; Version = $version; Frameworks = @() }
                } elseif ($packages.Count -gt 1) {
                    throw "Benchmark library '$($library.id)' lists several packages; add 'declare' with the package and version the comparisons measure."
                } elseif ($versions[0] -notmatch '^[^\[\(\*, ]+$') {
                    # Not a pinned version, such as "2026.6.0 or later"; such a library cannot be declared.
                    continue
                } else {
                    [pscustomobject]@{ Id = [string] $library.name; Package = [string] $packages[0].id; Version = $versions[0]; Frameworks = @() }
                }
            }
        )
        foreach ($item in $declared) {
            $names = @(
                $item.Id
                if ($declared.Count -eq 1) { [string] $library.id; [string] $library.name; @($library.aliases | Where-Object { $_ }) }
            ) | Sort-Object -Unique
            [pscustomobject]@{
                Id = $item.Id
                PolicyId = [string] $library.id
                Names = @($names)
                Package = $item.Package
                Version = $item.Version
                Frameworks = @($item.Frameworks)
                License = [string] $library.license.expression
                LicenseUrl = [string] $library.license.url
                Restriction = [string] $library.restriction
                Reason = [string] $library.reason
            }
        }
    }
}

function Resolve-BenchmarkDeclaredLibrary {
    <#
    .SYNOPSIS
    Finds the declared library for a name used in a comparison's Libraries.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [object[]] $Libraries,
        [Parameter(Mandatory)] [string] $Name
    )

    $found = @($Libraries | Where-Object { $Name -in $_.Names })
    if ($found.Count -ne 1) {
        throw "Comparison library '$Name' must resolve to exactly one declared library; it resolves to $($found.Count)."
    }
    $found[0]
}

function ConvertTo-BenchmarkIdentityKey {
    param([string] $Value)
    ($Value -replace '[^A-Za-z0-9]', '').ToLowerInvariant()
}

function Get-BenchmarkMethodIdentity {
    <#
    .SYNOPSIS
    Resolves the library and the shared case a BenchmarkDotNet method of a comparison measures.
    .DESCRIPTION
    The method name starts with its library ("MagickNETEncode" is Magick.NET's "Encode" case, "Sylvan" is the
    Sylvan library's default case). A comparison of one library uses the method name as its case. Methods that do
    not follow the pattern are listed in the definition's MethodIdentity ("Method" = "Library|Case") or in
    $BenchmarkMethodIdentityOverrides. Ambiguous or unmatched methods fail, so
    the importer never has to guess.
    .OUTPUTS
    An object with Library (declared id) and Case.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Workload,
        [Parameter(Mandatory)] [object] $Definition,
        [Parameter(Mandatory)] [object[]] $Libraries,
        [Parameter(Mandatory)] [string] $Method
    )

    $comparisonLibraries = @($Definition.Libraries | ForEach-Object { Resolve-BenchmarkDeclaredLibrary -Libraries $Libraries -Name $_ })
    $override = $script:BenchmarkMethodIdentityOverrides["$Workload|$Method"]
    if ($null -ne $Definition.PSObject.Properties['MethodIdentity'] -and $Definition.MethodIdentity.Contains($Method)) {
        $override = [string] $Definition.MethodIdentity[$Method]
    }
    if ($null -ne $override) {
        $libraryName, $case = $override -split '\|', 2
        return [pscustomobject]@{ Library = (Resolve-BenchmarkDeclaredLibrary -Libraries $Libraries -Name $libraryName).Id; Case = $case }
    }
    if ($comparisonLibraries.Count -eq 1) {
        return [pscustomobject]@{ Library = $comparisonLibraries[0].Id; Case = $Method }
    }

    $methodKey = ConvertTo-BenchmarkIdentityKey $Method
    $leading = ConvertTo-BenchmarkIdentityKey (($Method -split '_', 2)[0])
    $candidates = @(
        foreach ($library in $comparisonLibraries) {
            foreach ($name in $library.Names) {
                $key = ConvertTo-BenchmarkIdentityKey $name
                if ($key.Length -eq 0) { continue }
                if ($methodKey.StartsWith($key, [StringComparison]::Ordinal)) {
                    [pscustomobject]@{ Library = $library; Length = $key.Length; Exact = $true }
                } elseif ($leading.Length -ge 4 -and $key.StartsWith($leading, [StringComparison]::Ordinal)) {
                    # The method abbreviates the library ("HypeLab" for HypeLab.IO.Excel); the whole leading segment names it.
                    [pscustomobject]@{ Library = $library; Length = $leading.Length; Exact = $false }
                }
            }
        }
    )
    if ($candidates.Count -eq 0) {
        throw "Method '$Method' of workload '$Workload' does not name one of its libraries ($(@($Definition.Libraries) -join ', '))."
    }
    $best = @($candidates | Sort-Object -Property @{ Expression = 'Length'; Descending = $true }, @{ Expression = 'Exact'; Descending = $true })
    $top = @($best | Where-Object { $_.Length -eq $best[0].Length -and $_.Exact -eq $best[0].Exact } | ForEach-Object { $_.Library.Id } | Sort-Object -Unique)
    if ($top.Count -ne 1) {
        throw "Method '$Method' of workload '$Workload' matches several libraries: $($top -join ', ')."
    }
    $library = $best[0].Library

    # The case is what follows the library's part of the method name, in the method's own spelling.
    $consumed = 0
    $remaining = $best[0].Length
    while ($remaining -gt 0 -and $consumed -lt $Method.Length) {
        if ($Method[$consumed] -match '[A-Za-z0-9]') { $remaining-- }
        $consumed++
    }
    if (-not $best[0].Exact) {
        $consumed = ($Method -split '_', 2)[0].Length
    }
    $case = $Method.Substring($consumed).TrimStart('_')
    # "OfficeIMOYaml" and "OfficeIMOToml" name the format after the library; the format is not a case.
    $areaToken = (Get-Culture).TextInfo.ToTitleCase([string] $Definition.Area)
    if ($case.StartsWith($areaToken, [StringComparison]::Ordinal)) {
        $case = $case.Substring($areaToken.Length).TrimStart('_')
    }
    if ([string]::IsNullOrWhiteSpace($case)) { $case = 'Default' }
    [pscustomobject]@{ Library = $library.Id; Case = $case }
}

function ConvertFrom-BenchmarkExpectedCase {
    <#
    .SYNOPSIS
    Splits a runner expected-case string ("Method|Name=Value&Name=Value") into its method and parameters.
    #>
    param([Parameter(Mandatory)] [string] $ExpectedCase)

    $method, $variables = $ExpectedCase -split '\|', 2
    $parameters = [ordered]@{}
    if (-not [string]::IsNullOrWhiteSpace($variables)) {
        foreach ($part in $variables -split '&') {
            $name, $value = $part -split '=', 2
            $parameters[$name] = $value
        }
    }
    [pscustomobject]@{ Method = $method; Parameters = $parameters }
}

function Get-BenchmarkComparisonImportMaps {
    <#
    .SYNOPSIS
    Returns the library and case maps Import-BenchmarkResult needs for one workload's methods.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Workload,
        [Parameter(Mandatory)] [object] $Definition,
        [Parameter(Mandatory)] [object[]] $Libraries
    )

    $libraryMap = @{}
    $caseMap = @{}
    foreach ($method in @($Definition.ExpectedCases | ForEach-Object { (ConvertFrom-BenchmarkExpectedCase $_).Method } | Select-Object -Unique)) {
        $identity = Get-BenchmarkMethodIdentity -Workload $Workload -Definition $Definition -Libraries $Libraries -Method $method
        $libraryMap[$method] = $identity.Library
        if ($identity.Case -cne $method) { $caseMap[$method] = $identity.Case }
    }
    [pscustomobject]@{ LibraryMap = $libraryMap; CaseMap = $caseMap }
}

function Get-BenchmarkSparseCells {
    <#
    .SYNOPSIS
    Lists a comparison's case x library x axis-value cells when the benchmark code does not define all of them.
    .DESCRIPTION
    Without cells, PowerForge expects every combination of the declared cases, libraries and axis values. When the
    expected cases cover all of them, nothing is returned. Otherwise each defined cell is returned once; a cell that
    every library of the comparison defines omits its library.
    .OUTPUTS
    Ordered dictionaries with case, an optional library and parameters, ready for the declaration document.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [AllowEmptyCollection()] [System.Collections.Generic.List[object]] $Cells,
        [Parameter(Mandatory)] [string[]] $Libraries,
        [Parameter(Mandatory)] [System.Collections.Generic.List[string]] $Cases,
        [Parameter(Mandatory)] [System.Collections.Specialized.OrderedDictionary] $Axes
    )

    $groups = [ordered]@{}
    foreach ($cell in $Cells) {
        $parts = @(
            foreach ($axis in $Axes.Keys) {
                if (-not $cell.Parameters.Contains($axis)) {
                    throw "Case '$($cell.Case)' of library '$($cell.Library)' gives no value for axis '$axis'; every case must set every axis."
                }
                '{0}={1}' -f $axis, $cell.Parameters[$axis]
            }
        )
        $key = '{0}|{1}' -f $cell.Case, ($parts -join '&')
        if (-not $groups.Contains($key)) {
            $groups[$key] = [pscustomobject]@{ Case = $cell.Case; Parameters = $cell.Parameters; Libraries = [System.Collections.Generic.List[string]]::new() }
        }
        if (-not $groups[$key].Libraries.Contains($cell.Library)) { $groups[$key].Libraries.Add($cell.Library) }
    }

    $complete = $Cases.Count * $Libraries.Count
    foreach ($values in $Axes.Values) { $complete *= $values.Count }
    $defined = 0
    foreach ($group in $groups.Values) { $defined += $group.Libraries.Count }
    if ($defined -eq $complete) { return }

    foreach ($group in $groups.Values) {
        $parameters = [ordered]@{}
        foreach ($axis in $Axes.Keys) { $parameters[$axis] = [string] $group.Parameters[$axis] }
        $everyLibrary = @($Libraries | Where-Object { $_ -notin $group.Libraries }).Count -eq 0
        $cellLibraries = if ($everyLibrary) { @('') } else { @($Libraries | Where-Object { $_ -in $group.Libraries }) }
        foreach ($library in $cellLibraries) {
            $cell = [ordered]@{ case = $group.Case }
            if ($library) { $cell.library = $library }
            if ($parameters.Count -gt 0) { $cell.parameters = $parameters }
            $cell
        }
    }
}

function New-BenchmarkDeclarationDocument {
    <#
    .SYNOPSIS
    Builds the schema 5 declaration document for every runner workload.
    .DESCRIPTION
    Cases and axes come from the workload's expected cases, libraries from its Libraries and the library policy.
    A comparison declares a framework only when every library runs its declared version there and the
    framework's cases match; otherwise lanes on that framework could not be published.
    .OUTPUTS
    An ordered dictionary ready for JSON serialization.
    #>
    [CmdletBinding()]
    param([object] $Policy = (Get-BenchmarkLibraryPolicy))

    $libraries = @(Get-BenchmarkDeclaredLibraries -Policy $Policy)
    $definitionsByFramework = @{}
    foreach ($framework in $script:BenchmarkDeclarationFrameworks) {
        $definitionsByFramework[$framework] = Get-LibraryComparisonDefinitions -Framework $framework
    }
    $primary = $definitionsByFramework['net10.0']
    $usedLibraries = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $comparisons = foreach ($workload in $primary.Keys) {
        $definition = $primary[$workload]
        $comparisonLibraries = @($definition.Libraries | ForEach-Object { Resolve-BenchmarkDeclaredLibrary -Libraries $libraries -Name $_ })
        foreach ($library in $comparisonLibraries) { $null = $usedLibraries.Add($library.Id) }

        $cases = [System.Collections.Generic.List[string]]::new()
        $axes = [ordered]@{}
        $expectedCells = [System.Collections.Generic.List[object]]::new()
        foreach ($expected in @($definition.ExpectedCases)) {
            $parsed = ConvertFrom-BenchmarkExpectedCase $expected
            $identity = Get-BenchmarkMethodIdentity -Workload $workload -Definition $definition -Libraries $libraries -Method $parsed.Method
            if (-not $cases.Contains($identity.Case)) { $cases.Add($identity.Case) }
            foreach ($parameter in $parsed.Parameters.GetEnumerator()) {
                if (-not $axes.Contains($parameter.Key)) { $axes[$parameter.Key] = [System.Collections.Generic.List[string]]::new() }
                if (-not $axes[$parameter.Key].Contains([string] $parameter.Value)) { $axes[$parameter.Key].Add([string] $parameter.Value) }
            }
            $expectedCells.Add([pscustomobject]@{ Library = $identity.Library; Case = $identity.Case; Parameters = $parsed.Parameters })
        }
        $cells = @(Get-BenchmarkSparseCells -Cells $expectedCells -Libraries @($comparisonLibraries.Id) -Cases $cases -Axes $axes)

        $frameworks = @(
            foreach ($framework in $script:BenchmarkDeclarationFrameworks) {
                if ($definition.Frameworks -and $framework -notin @($definition.Frameworks)) { continue }
                if (@($comparisonLibraries | Where-Object { $_.Frameworks.Count -gt 0 -and $framework -notin $_.Frameworks }).Count -gt 0) { continue }
                $frameworkCases = @($definitionsByFramework[$framework][$workload].ExpectedCases | Sort-Object)
                if (($frameworkCases -join "`n") -cne (@($definition.ExpectedCases | Sort-Object) -join "`n")) { continue }
                $framework
            }
        )
        if ($frameworks.Count -eq 0) {
            throw "Workload '$workload' has no framework on which every library runs its declared version."
        }

        $comparison = [ordered]@{
            comparisonId = [string] $definition.ComparisonId
            area = [string] $definition.Area
            title = [string] $definition.Title
            subject = (Resolve-BenchmarkDeclaredLibrary -Libraries $libraries -Name 'OfficeIMO').Id
            libraries = @($comparisonLibraries.Id)
            axes = @(foreach ($axis in $axes.GetEnumerator()) { [ordered]@{ name = $axis.Key; values = @($axis.Value) } })
            cases = @($cases)
        }
        if ($cells.Count -gt 0) { $comparison.cells = $cells }
        $comparison.expectedPlatforms = @($definition.ExpectedPlatforms)
        $comparison.frameworks = $frameworks
        if (-not $definition.CatalogEligible) {
            $reason = [string] $definition.PublishRestriction
            if ([string]::IsNullOrWhiteSpace($reason) -or $reason -eq 'None') { $reason = 'Not a public library comparison.' }
            $comparison.eligibility = [ordered]@{ publishable = $false; reason = $reason }
        }
        $comparison
    }

    [ordered]@{
        '$schema' = 'https://raw.githubusercontent.com/EvotecIT/PSPublishModule/main/Schemas/powerforge.benchmark-declarations.schema.json'
        schemaVersion = 5
        libraries = @(
            foreach ($library in $libraries | Where-Object { $usedLibraries.Contains($_.Id) } | Sort-Object Id) {
                $entry = [ordered]@{
                    id = $library.Id
                    package = $library.Package
                    version = $library.Version
                }
                if ($library.License) { $entry.license = $library.License }
                if ($library.LicenseUrl) { $entry.licenseUrl = $library.LicenseUrl }
                $entry.restriction = $library.Restriction
                if ($library.Restriction -ne 'open') { $entry.reason = $library.Reason }
                $entry
            }
        )
        comparisons = @($comparisons)
    }
}
