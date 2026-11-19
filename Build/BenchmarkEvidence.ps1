# The method and identity parameters of an imported row, in the runner's expected-case form
# ("Method|Name=Value&Name=Value", names sorted). A row whose case was renamed keeps its method in the Method variable.
function Get-BenchmarkMethodCombination {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object] $Row,
        [AllowEmptyCollection()]
        [string[]] $VariableName = @()
    )

    $method = [string] $Row.Scenario
    $parts = [System.Collections.Generic.List[object]]::new()
    if ($null -ne $Row.Variables) {
        foreach ($variable in $Row.Variables.GetEnumerator()) {
            $key = [string] $variable.Key
            if ($key -ieq 'Method') {
                if (-not [string]::IsNullOrWhiteSpace([string] $variable.Value)) { $method = [string] $variable.Value }
                continue
            }
            # BenchmarkDotNet can report several parameters in one value ("A=1&B=2").
            foreach ($part in [regex]::Split(('{0}={1}' -f $key, [string] $variable.Value), '&(?=[^=&]+=)')) {
                $equalsIndex = $part.IndexOf('=')
                if ($equalsIndex -lt 1) { continue }
                $name = $part.Substring(0, $equalsIndex)
                if ($name.ToLowerInvariant() -in @($VariableName | ForEach-Object { $_.ToLowerInvariant() })) {
                    $parts.Add([pscustomobject]@{ Key = $name; Value = $part.Substring($equalsIndex + 1) })
                }
            }
        }
    }
    if ([string]::IsNullOrWhiteSpace($method)) {
        throw 'Benchmark rows need a method identity.'
    }
    $variables = @($parts | Sort-Object { [string] $_.Key } | ForEach-Object { '{0}={1}' -f $_.Key, $_.Value })
    if ($variables.Count -eq 0) { return $method }
    '{0}|{1}' -f $method, ($variables -join '&')
}

function Write-BenchmarkEvidenceResult {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string] $Path,
        [Parameter(Mandatory)]
        [object] $InputObject
    )

    $benchmarkJsonType = [Type]::GetType('PowerForge.BenchmarkJson, PowerForge', $true)
    $writeMethod = $benchmarkJsonType.GetMethods() |
        Where-Object { $_.Name -eq 'Write' -and $_.IsGenericMethodDefinition } |
        Select-Object -First 1
    if ($null -eq $writeMethod) {
        throw 'PowerForge.BenchmarkJson.Write<T> is unavailable.'
    }

    $value = $InputObject.PSObject.BaseObject
    $closedMethod = $writeMethod.MakeGenericMethod($value.GetType())
    $arguments = [object[]] @([string] $Path, $value)
    $null = $closedMethod.Invoke($null, $arguments)
}

# BenchmarkDotNet job of the standard profile: one launch, 5 warmups and 15 measured iterations per case, keeping
# every sample. The publish policy needs at least 3 warmups and 10 samples for a confidence interval of the median.
function Get-StandardBenchmarkJobArguments {
    @('--launchCount', '1', '--warmupCount', '5', '--iterationCount', '15', '--outliers', 'DontRemove')
}

# Version and SHA-256 of the restored package of every compared library, as publish-policy metadata
# (benchmark.package.<package> and benchmark.package.<package>.sha256). The version comes from the project's restore
# result, so a library that resolved to another version than the declared one is recorded as it ran.
function Get-BenchmarkPackageProvenance {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string] $AssetsPath,
        [Parameter(Mandatory)]
        [string] $Framework,
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [object[]] $Libraries
    )

    if (-not (Test-Path -LiteralPath $AssetsPath -PathType Leaf)) {
        throw "The restore result '$AssetsPath' was not found; restore the benchmark project first."
    }
    $assets = Get-Content -LiteralPath $AssetsPath -Raw -Encoding UTF8 | ConvertFrom-Json -AsHashtable
    $packageFolders = @($assets['packageFolders'].Keys)
    # The packages restored for this framework only; other frameworks of the project may resolve other versions.
    $targetKey = @($assets['targets'].Keys | Where-Object { $_ -eq $Framework -or $_.StartsWith("$Framework/", [StringComparison]::Ordinal) } | Select-Object -First 1)
    if ($targetKey.Count -eq 0) {
        throw "The restore result '$AssetsPath' has no $Framework target."
    }
    $target = $assets['targets'][$targetKey[0]]
    $metadata = [ordered]@{}
    foreach ($library in $Libraries) {
        $package = [string] $library.Package
        $restored = @($target.Keys | Where-Object {
                ($_ -split '/', 2)[0] -ieq $package -and $target[$_]['type'] -eq 'package'
            })
        if ($restored.Count -ne 1) { continue }
        $version = ($restored[0] -split '/', 2)[1]
        $key = 'benchmark.package.' + $package.ToLowerInvariant()
        $metadata[$key] = $version
        $relative = [string] $assets['libraries'][$restored[0]]['path']
        $nupkg = $null
        foreach ($folder in $packageFolders) {
            $candidate = Join-Path (Join-Path $folder $relative) ('{0}.{1}.nupkg' -f $package.ToLowerInvariant(), $version.ToLowerInvariant())
            if (Test-Path -LiteralPath $candidate -PathType Leaf) { $nupkg = $candidate; break }
        }
        if ($null -ne $nupkg) {
            $metadata[$key + '.sha256'] = (Get-FileHash -LiteralPath $nupkg -Algorithm SHA256).Hash.ToLowerInvariant()
        }
    }
    $metadata
}

# Creates the same argv arrays used by execution and side-effect-free planning.
function Get-LibraryComparisonCommands {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object] $Definition,
        [Parameter(Mandatory)]
        [string] $RepositoryRoot,
        [Parameter(Mandatory)]
        [string] $Framework,
        [Parameter(Mandatory)]
        [ValidateSet('diagnostic', 'standard', 'full')]
        [string] $BenchmarkProfile,
        [Parameter(Mandatory)]
        [string] $ArtifactsPath,
        [string] $BuildArtifactsPath,
        [bool] $KeepBenchmarkFiles,
        # MSBuild properties that the caller's accepted library licences require (from the library policy).
        [System.Collections.IDictionary] $BuildProperties = @{},
        [string] $QuestPdfVersion
    )

    $hostArguments = @(
        'run', '-c', 'Release', '-f', $Framework,
        '--project', (Join-Path $RepositoryRoot $Definition.Project)
    )
    if (-not [string]::IsNullOrWhiteSpace($BuildArtifactsPath)) {
        $hostArguments += @('--artifacts-path', [System.IO.Path]::GetFullPath($BuildArtifactsPath))
    }
    $buildArguments = @()
    if (-not [string]::IsNullOrWhiteSpace($QuestPdfVersion)) {
        $buildArguments += "-p:QuestPdfBenchmarkVersion=$QuestPdfVersion"
    }
    foreach ($property in $BuildProperties.GetEnumerator()) {
        $buildArguments += "-p:$($property.Key)=$($property.Value)"
    }
    if ($Definition.Suite -in @('OfficeIMO.Pdf.HtmlGeneration', 'OfficeIMO.Pdf.Html21KiBPayload')) {
        $packagePath = Join-Path (Split-Path -Parent $ArtifactsPath) (
            '{0}-pdfpecker-package.json' -f (Split-Path -Leaf $ArtifactsPath))
        [pscustomobject]@{
            Kind = 'package-evidence'
            FilePath = 'dotnet'
            Arguments = @($hostArguments + @('--', 'html-package-evidence', '--output', $packagePath))
        }
    }
    $runnerArguments = @(if ($null -ne $Definition.RunnerArguments) { $Definition.RunnerArguments })
    $arguments = $hostArguments + $buildArguments + @('--') + $runnerArguments + @(
        '--filter', $Definition.Filter, '--artifacts', $ArtifactsPath,
        '--exporters', 'fulljson'
    )
    if ($Definition.Project -like '*Pdf.Benchmarks.Comparisons*') {
        $arguments += @('--buildTimeout', '600')
    }
    if ($KeepBenchmarkFiles) { $arguments += @('--keepFiles', '--generateBinLog') }
    # diagnostic: one BenchmarkDotNet dry pass, never published. standard: enough warmups and samples for the publish
    # policy (median with a confidence interval) in far less time than a full run. full: the workload's own job.
    if ($BenchmarkProfile -eq 'diagnostic') { $arguments += @('--job', 'Dry') }
    elseif ($BenchmarkProfile -eq 'standard') { $arguments += Get-StandardBenchmarkJobArguments }
    elseif ($null -ne $Definition.PSObject.Properties['FullRunArguments']) {
        $arguments += $Definition.FullRunArguments
    }
    [pscustomobject]@{
        Kind = 'measure'
        FilePath = 'dotnet'
        Arguments = $arguments
    }
    if ($Definition.ValidatedSizeEvidence) {
        [pscustomobject]@{
            Kind = 'validate-size'
            FilePath = 'dotnet'
            Arguments = @($hostArguments + $buildArguments + @(
                '--', 'validate', '--json', (Join-Path $ArtifactsPath 'validated-size-evidence.json')
            ))
        }
    }
}

# Captures recheck source cleanliness for every workload, so in-tree output must stay ignored.
function Assert-BenchmarkOutputIgnored {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string] $RepositoryRoot,
        [Parameter(Mandatory)]
        [string[]] $Path
    )

    $sourceRoot = (& git -C $RepositoryRoot rev-parse --show-toplevel).Trim()
    if ($LASTEXITCODE -ne 0) { throw 'Unable to resolve source checkout for output containment.' }
    $comparison = if ($IsWindows) { [StringComparison]::OrdinalIgnoreCase } else { [StringComparison]::Ordinal }
    foreach ($outputPath in $Path) {
        $fullPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($outputPath)
        # Git resolves existing parent aliases, including the configured repository-root symlinks.
        $existingParent = $fullPath
        while (-not (Test-Path -LiteralPath $existingParent -PathType Container)) {
            $existingParent = Split-Path -Parent $existingParent
        }
        $outputCheckout = & git -C $existingParent rev-parse --show-toplevel 2>$null
        if ($LASTEXITCODE -ne 0 -or -not ([string] $outputCheckout).Trim().Equals($sourceRoot, $comparison)) {
            continue
        }
        $prefix = [string] (& git -C $existingParent rev-parse --show-prefix)
        if ($LASTEXITCODE -ne 0) { throw "Unable to resolve benchmark output '$fullPath'." }
        $relativePath = [System.IO.Path]::GetRelativePath($existingParent, $fullPath)
        $probe = $prefix + $relativePath.Replace('\', '/') + '/.officeimo-benchmark-output-' + [guid]::NewGuid().ToString('N')
        & git -C $sourceRoot check-ignore --quiet -- $probe
        if ($LASTEXITCODE -ne 0) {
            throw "Benchmark output '$fullPath' is inside the source checkout and is not gitignored. Use the default temporary bundle, an external scratch path, or an ignored path such as artifacts/benchmark-bundles."
        }
    }
}
