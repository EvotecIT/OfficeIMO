# Reads Build/benchmark-libraries.json and resolves which libraries a benchmark run may use. Loading this file performs no work.

# Benchmark code reads this variable: a comma-separated list of the restricted library ids the caller enabled.
$script:BenchmarkEnabledLibrariesVariable = 'OFFICEIMO_BENCHMARK_ENABLED_LIBRARIES'

function Get-BenchmarkLibraryPolicy {
    <#
    .SYNOPSIS
    Loads the benchmark library policy.
    .PARAMETER Path
    Policy file; defaults to Build/benchmark-libraries.json.
    .OUTPUTS
    The policy object with a Libraries array.
    #>
    [CmdletBinding()]
    param([string] $Path = (Join-Path $PSScriptRoot 'benchmark-libraries.json'))

    $policy = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json -Depth 20
    $classes = @($policy.restrictionClasses.PSObject.Properties.Name)
    $seen = @{}
    foreach ($library in @($policy.libraries)) {
        if ([string] $library.restriction -notin $classes) {
            throw "Benchmark library '$($library.id)' has unknown restriction '$($library.restriction)'."
        }
        foreach ($key in @(@([string] $library.id, [string] $library.name) + @($library.aliases | Where-Object { $_ }) | Sort-Object -Unique)) {
            $normalized = $key.ToLowerInvariant()
            if ($seen.ContainsKey($normalized) -and $seen[$normalized] -ne $library.id) {
                throw "Benchmark library name '$key' is declared by both '$($seen[$normalized])' and '$($library.id)'."
            }
            $seen[$normalized] = $library.id
        }
    }
    $policy
}

function Resolve-BenchmarkLibrary {
    <#
    .SYNOPSIS
    Finds a library by id, name or alias.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [object] $Policy,
        [Parameter(Mandatory)] [string] $Name
    )

    $match = @($Policy.libraries | Where-Object {
            $_.id -eq $Name -or $_.name -eq $Name -or ($_.aliases -and $Name -in @($_.aliases))
        }) | Select-Object -First 1
    if ($null -eq $match) {
        throw "Benchmark library '$Name' is not declared in Build/benchmark-libraries.json."
    }
    $match
}

function Get-BenchmarkLibrarySelection {
    <#
    .SYNOPSIS
    Applies -IncludeLibrary, -ExcludeLibrary and -AcceptLibraryLicense to the policy.
    .DESCRIPTION
    Open libraries run unless excluded. Publish-restricted libraries run only when included. Acceptance-required
    libraries run only when their licence is accepted. Excluded libraries never run. Nothing here declares a
    licence: the caller supplies whatever the library itself reads.
    .OUTPUTS
    Enabled: ids of restricted libraries enabled for this run. Excluded: ids removed with -ExcludeLibrary.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [object] $Policy,
        [string[]] $IncludeLibrary = @(),
        [string[]] $ExcludeLibrary = @(),
        [string[]] $AcceptLibraryLicense = @()
    )

    $split = { param([string[]] $Values) @($Values | ForEach-Object { ([string] $_) -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ }) }
    $included = @(& $split $IncludeLibrary | ForEach-Object { (Resolve-BenchmarkLibrary -Policy $Policy -Name $_).id })
    $excluded = @(& $split $ExcludeLibrary | ForEach-Object { (Resolve-BenchmarkLibrary -Policy $Policy -Name $_).id })
    $accepted = @(& $split $AcceptLibraryLicense | ForEach-Object { (Resolve-BenchmarkLibrary -Policy $Policy -Name $_).id })

    $enabled = [System.Collections.Generic.List[string]]::new()
    foreach ($id in @($included + $accepted | Select-Object -Unique)) {
        if ($id -in $excluded) { continue }
        $library = Resolve-BenchmarkLibrary -Policy $Policy -Name $id
        switch ([string] $library.restriction) {
            'excluded' { throw "Benchmark library '$id' is excluded by the library policy: $($library.reason)" }
            'open' { continue }
            default {
                if ($null -ne $library.acceptance -and $id -notin $accepted) {
                    throw "Benchmark library '$id' needs -AcceptLibraryLicense $id after you review its licence ($($library.license.url))."
                }
                $enabled.Add($id)
            }
        }
    }

    [pscustomobject]@{
        Enabled = @($enabled | Sort-Object -Unique)
        Excluded = @($excluded | Sort-Object -Unique)
    }
}

function Test-BenchmarkCaseMatchesLibrary {
    <#
    .SYNOPSIS
    Returns whether a case identity (method name, optionally followed by |variables) belongs to a library.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [object] $Library,
        [Parameter(Mandatory)] [string] $Case
    )

    $method = ($Case -split '\|', 2)[0]
    foreach ($pattern in @($Library.cases | Where-Object { $_ })) {
        if ($method -like $pattern) { return $true }
    }
    $false
}

function Get-BenchmarkLibraryAcceptanceIssues {
    <#
    .SYNOPSIS
    Lists the values an enabled library needs from the caller that are missing in this process.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)] [object] $Library)

    foreach ($variable in @($Library.acceptance.environmentVariables | Where-Object { $_ })) {
        if ([string]::IsNullOrWhiteSpace([Environment]::GetEnvironmentVariable($variable))) {
            "$($Library.name) needs your own licence declaration in the $variable environment variable."
        }
    }
}
