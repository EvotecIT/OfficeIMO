function Get-BenchmarkEvidenceLocation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string] $ComparisonId,
        [Parameter(Mandatory)]
        [ValidateSet('windows', 'linux', 'macos')]
        [string] $Platform,
        [Parameter(Mandatory)]
        [ValidateSet('quick', 'full')]
        [string] $RunMode,
        [Parameter(Mandatory)]
        [string] $StaticRoot,
        [switch] $Portable
    )

    $comparisonSlug = ($ComparisonId.Trim() -replace '[^A-Za-z0-9._-]', '-').ToLowerInvariant()
    if ([string]::IsNullOrWhiteSpace($comparisonSlug)) {
        throw 'ComparisonId must contain at least one filename-safe character.'
    }

    $fileName = "$comparisonSlug-$Platform-$RunMode.json"
    [pscustomobject]@{
        FileName = $fileName
        Path = Join-Path $StaticRoot $fileName
        ResultPath = if ($Portable) { $fileName } else { "/data/benchmarks/library-comparisons/$fileName" }
    }
}

function Get-BenchmarkEvidenceCaseIdentity {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object] $Row,
        [string[]] $VariableName
    )

    $scenario = [string] $Row.Scenario
    if ([string]::IsNullOrWhiteSpace($scenario)) {
        throw 'Benchmark evidence rows require a scenario identity.'
    }

    $ignoredVariables = @('namespace', 'type', 'fullname')
    $restrictVariables = $PSBoundParameters.ContainsKey('VariableName')
    $variables = @(
        if ($null -ne $Row.Variables) {
            foreach ($variable in $Row.Variables.GetEnumerator()) {
                $combined = '{0}={1}' -f ([string] $variable.Key), ([string] $variable.Value)
                foreach ($part in [regex]::Split($combined, '&(?=[^=&]+=)')) {
                    $equalsIndex = $part.IndexOf('=')
                    if ($equalsIndex -lt 1) {
                        continue
                    }

                    [pscustomobject]@{
                        Key = $part.Substring(0, $equalsIndex)
                        Value = $part.Substring($equalsIndex + 1)
                    }
                }
            }
        }
    )
    $variableParts = @(
        $variables |
            Where-Object {
                $key = ([string] $_.Key).ToLowerInvariant()
                $key -notin $ignoredVariables -and
                (-not $restrictVariables -or $key -in $VariableName)
            } |
            Sort-Object { [string] $_.Key } |
            ForEach-Object {
                '{0}={1}' -f ([string] $_.Key), ([string] $_.Value)
            }
    )

    if ($variableParts.Count -eq 0) {
        return $scenario
    }

    return '{0}|{1}' -f $scenario, ($variableParts -join '&')
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
        [string] $RunMode,
        [Parameter(Mandatory)]
        [string] $ArtifactsPath,
        [string] $BuildArtifactsPath,
        [bool] $KeepBenchmarkFiles,
        [bool] $Word,
        [string] $QuestPdfVersion,
        [bool] $InternalQuestPdf
    )

    $hostArguments = @(
        'run', '-c', 'Release', '-f', $Framework,
        '--project', (Join-Path $RepositoryRoot $Definition.Project)
    )
    if (-not [string]::IsNullOrWhiteSpace($BuildArtifactsPath)) {
        $hostArguments += @('--artifacts-path', [System.IO.Path]::GetFullPath($BuildArtifactsPath))
    }
    $buildArguments = @()
    if ($Word) { $buildArguments += '-p:AcceptNPOIOSMFLicense=true' }
    if (-not [string]::IsNullOrWhiteSpace($QuestPdfVersion)) {
        $buildArguments += "-p:QuestPdfBenchmarkVersion=$QuestPdfVersion"
        $buildArguments += "-p:QuestPdfInternalAuthorization=$($InternalQuestPdf.ToString().ToLowerInvariant())"
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
        '--exporters', 'jsonfull'
    )
    if ($Definition.Project -like '*Pdf.Benchmarks.Comparisons*') {
        $arguments += @('--buildTimeout', '600')
    }
    if ($KeepBenchmarkFiles) { $arguments += @('--keepFiles', '--generateBinLog') }
    if ($RunMode -eq 'quick') { $arguments += @('--job', 'Dry') }
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
