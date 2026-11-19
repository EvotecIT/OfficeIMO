param(
    # diagnostic: a quick check, never published. standard: enough samples for the publish policy.
    # full: the complete BenchmarkDotNet job of each workload.
    [ValidateSet('diagnostic', 'standard', 'full')]
    [Alias('Profile', 'RunMode')]
    [string] $BenchmarkProfile = 'full',
    [ValidateSet('net8.0', 'net10.0')]
    [string] $Framework = 'net10.0',
    [string[]] $Workload = @('all'),
    [ValidateSet('windows', 'linux', 'macos')]
    [string] $Platform,
    [Alias('BundleRoot')]
    [string] $EvidenceRoot,
    [switch] $DryRun,
    [string] $PdfCorpusRoot,
    [string] $HtmlTinkerXRoot = $env:HTMLTINKERX_ROOT,
    [string] $OutputRoot = (Join-Path ([System.IO.Path]::GetTempPath()) 'OfficeIMO/Benchmarks/Runs'),
    [string] $PowerForgeRoot = $env:POWERFORGE_ROOT,
    [string] $BuildArtifactsPath,
    [UInt64] $AffinityMask = 0,
    # A QuestPDF release other than the pinned public one; it resolves to the 'questpdf-latest' library policy.
    [string] $QuestPdfPackageVersion,
    # Restricted libraries to run (ids or names from Build/benchmark-libraries.json).
    [string[]] $IncludeLibrary = @(),
    # Libraries to leave out; a comparison that needs one of them is skipped.
    [string[]] $ExcludeLibrary = @(),
    # Licences you have reviewed and accepted; the library's own key or setting still comes from you.
    [string[]] $AcceptLibraryLicense = @(),
    # Keeps every result out of the shared catalog, whatever libraries the run uses.
    [switch] $LocalOnly,
    [switch] $Publish,
    [switch] $PlanOnly,
    [switch] $KeepBenchmarkFiles,
    [ValidateSet('Default', 'WarmedV1')]
    [string] $MeasurementProfile = 'Default'
)

$ErrorActionPreference = 'Stop'

if ($Publish -and $BenchmarkProfile -eq 'diagnostic') {
    throw 'A diagnostic run is never published; use -Profile standard or full.'
}

if ($MeasurementProfile -eq 'WarmedV1' -and
    ($BenchmarkProfile -ne 'full' -or $Workload.Count -ne 1 -or $Workload[0] -ne 'pdfhtmlpayload')) {
    throw 'WarmedV1 requires -Profile full -Workload pdfhtmlpayload.'
}

. (Join-Path $PSScriptRoot 'BenchmarkEvidence.ps1')
. (Join-Path $PSScriptRoot 'BenchmarkDeclarations.ps1')
$declarationPath = Join-Path $PSScriptRoot 'benchmark-declarations.json'
$declarations = Get-Content -LiteralPath $declarationPath -Raw -Encoding UTF8 | ConvertFrom-Json

$repositoryRoot = Split-Path -Parent $PSScriptRoot
$questPdfBenchmarkVersion = '2026.5.0'
$questPdfWorkloads = @('pdfgenerate', 'pdfinvoice', 'pdfread', 'pdfreverse')
$htmlTinkerXProjectPath = $null
$htmlTinkerXSourceRoot = $null
$htmlTinkerXSourceCommit = $null
$htmlTinkerXSourceDirty = $false
$isPlanning = $PlanOnly -or $DryRun
if ($PSBoundParameters.ContainsKey('Platform') -and -not $isPlanning) {
    throw '-Platform is a planning override only; actual collection uses the host platform.'
}
$usesHtmlTinkerX = 'pdfhtml' -in $Workload -or 'all' -in $Workload
if ($usesHtmlTinkerX -and -not [string]::IsNullOrWhiteSpace($HtmlTinkerXRoot)) {
    $resolvedHtmlTinkerX = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($HtmlTinkerXRoot)
    $htmlTinkerXProjectPath = if ($resolvedHtmlTinkerX.EndsWith('.csproj', [StringComparison]::OrdinalIgnoreCase)) {
        $resolvedHtmlTinkerX
    } else {
        Join-Path $resolvedHtmlTinkerX 'Sources\HtmlTinkerX\HtmlTinkerX.csproj'
    }
}

$OutputRoot = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputRoot)
$hostPlatform = if ($IsWindows) { 'windows' } elseif ($IsLinux) { 'linux' } elseif ($IsMacOS) { 'macos' } else {
    throw 'Unsupported benchmark platform.'
}
if ([string]::IsNullOrWhiteSpace($Platform)) { $Platform = $hostPlatform }
$isWindowsBenchmarkPlatform = $Platform -eq 'windows'
$isLinuxBenchmarkPlatform = $Platform -eq 'linux'
$canInheritAffinity = $isWindowsBenchmarkPlatform -or $isLinuxBenchmarkPlatform
$affinityLabel = if ($AffinityMask -ne 0) { '0x{0:X}' -f $AffinityMask } else { $null }
$affinityApplication = if ($AffinityMask -ne 0 -and $canInheritAffinity) { 'inherited-parent-process' } else { $null }
$effectiveAffinityApplication = if ($null -ne $affinityApplication) { $affinityApplication } elseif ($AffinityMask -ne 0) {
    'unsupported'
} elseif ($Platform -eq 'macos') { 'no-control-macos' } else { 'os-default' }
if (-not $isPlanning -and $Platform -eq 'macos' -and $AffinityMask -ne 0) {
    throw 'macOS does not support this runner''s affinity control; use -AffinityMask 0.'
}
$stamp = [DateTimeOffset]::UtcNow.ToString('yyyyMMdd-HHmmss-fff')
$websiteRoot = Join-Path $repositoryRoot 'Website/static/data/benchmarks/evidence'
$staticRoot = if ([string]::IsNullOrWhiteSpace($EvidenceRoot)) {
    Join-Path ([System.IO.Path]::GetTempPath()) "OfficeIMO/Benchmarks/Bundles/$Platform"
} else {
    $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($EvidenceRoot)
}
$portableBundle = $staticRoot -ne $websiteRoot
if ($portableBundle -and -not $PSBoundParameters.ContainsKey('OutputRoot')) {
    $OutputRoot = Join-Path $staticRoot 'raw'
}
$catalogPath = Join-Path $staticRoot 'index.json'

$definitions = Get-LibraryComparisonDefinitions -Framework $Framework

if ($MeasurementProfile -eq 'WarmedV1') {
    $definitions.pdfhtmlpayload | Add-Member -NotePropertyName FullRunArguments -NotePropertyValue @(
        '--strategy', 'Throughput',
        '--launchCount', '1',
        '--iterationTime', '1000',
        '--minWarmupCount', '64',
        '--maxWarmupCount', '128',
        '--minIterationCount', '30',
        '--maxIterationCount', '60',
        '--outliers', 'DontRemove'
    )
}

$includeAll = 'all' -in $Workload
$explicitWorkloads = @(
    foreach ($requested in $Workload) {
        if ($requested -eq 'all') { continue }
        if ($definitions.Contains($requested)) { $requested; continue }
        # An area name selects every workload of that area.
        $areaWorkloads = @($definitions.Keys | Where-Object { $definitions[$_].Area -eq $requested })
        if ($areaWorkloads.Count -gt 0) { $areaWorkloads; continue }
        $areas = @($definitions.Values | ForEach-Object { $_.Area } | Sort-Object -Unique)
        throw "Unknown workload '$requested'. Choose all, an area ($($areas -join ', ')), or: $($definitions.Keys -join ', ')."
    }
)
$selected = @(@(if ($includeAll) { $definitions.Keys }; $explicitWorkloads) | Select-Object -Unique)
if ($selected.Count -eq 0) { throw 'Select at least one workload.' }
$containsQuestPdfWorkload = @($selected | Where-Object { $_ -in $questPdfWorkloads }).Count -gt 0
$alternateQuestPdf = -not [string]::IsNullOrWhiteSpace($QuestPdfPackageVersion) -and $QuestPdfPackageVersion -ne $questPdfBenchmarkVersion
$effectiveQuestPdfVersion = if ($alternateQuestPdf) { $QuestPdfPackageVersion } else { $questPdfBenchmarkVersion }

$libraryPolicy = Get-BenchmarkLibraryPolicy
$librarySelection = Get-BenchmarkLibrarySelection -Policy $libraryPolicy -IncludeLibrary $IncludeLibrary `
    -ExcludeLibrary $ExcludeLibrary -AcceptLibraryLicense $AcceptLibraryLicense
if ($alternateQuestPdf) {
    if (-not $containsQuestPdfWorkload -or @($selected | Where-Object { $_ -notin $questPdfWorkloads }).Count -gt 0) {
        throw 'An alternate QuestPDF release accepts only workloads that actually exercise QuestPDF.'
    }
    if ('questpdf-latest' -notin $librarySelection.Enabled) {
        throw "QuestPDF $QuestPdfPackageVersion is restricted; review its licence and add -AcceptLibraryLicense questpdf-latest."
    }
}
# Any enabled restricted library keeps the whole run local: it may add cases to otherwise published comparisons.
$localOnlyRun = $LocalOnly -or $alternateQuestPdf -or $librarySelection.Enabled.Count -gt 0
if ($LocalOnly -and $Publish) {
    throw '-LocalOnly cannot be combined with -Publish.'
}
if ($localOnlyRun -and $Publish) {
    throw "Restricted libraries ($(@($librarySelection.Enabled) -join ', ')) cannot be published; drop -Publish or the restricted libraries."
}

if ($containsQuestPdfWorkload) {
    [xml] $comparisonProject = Get-Content -LiteralPath (
        Join-Path $repositoryRoot 'OfficeIMO.Pdf.Benchmarks.Comparisons\OfficeIMO.Pdf.Benchmarks.Comparisons.csproj') -Raw
    $questPdfReference = @($comparisonProject.Project.ItemGroup.PackageReference) |
        Where-Object Include -eq 'QuestPDF' |
        Select-Object -First 1
    if ($null -eq $questPdfReference -or [string] $questPdfReference.Version -ne '$(QuestPdfBenchmarkVersion)') {
        throw "QuestPDF benchmark workloads require the last MIT-compatible package version $questPdfBenchmarkVersion."
    }
}

$executionPlan = @(
    foreach ($name in $selected) {
        $definition = $definitions[$name]
        # Library policy: restricted libraries make a comparison unpublishable, and their cases are left out
        # unless the caller enabled them.
        $workloadLibraries = @(
            foreach ($libraryName in @($definition.Libraries)) {
                $library = Resolve-BenchmarkLibrary -Policy $libraryPolicy -Name $libraryName
                if ($alternateQuestPdf -and $library.id -eq 'questpdf') {
                    $library = Resolve-BenchmarkLibrary -Policy $libraryPolicy -Name 'questpdf-latest'
                }
                $library
            }
            # Every QuestPDF workload builds against the selected release, even when QuestPDF is not a listed peer.
            if ($alternateQuestPdf -and $name -in $questPdfWorkloads -and 'QuestPDF' -notin @($definition.Libraries)) {
                Resolve-BenchmarkLibrary -Policy $libraryPolicy -Name 'questpdf-latest'
            }
        )
        # A published lane drops restricted libraries' cases and publishes the rest. Only a restricted library whose
        # cases cannot be identified (it has no case patterns) makes the comparison unpublishable.
        $restrictedLibraries = @($workloadLibraries | Where-Object restriction -ne 'open')
        $inseparableLibraries = @($restrictedLibraries | Where-Object { @($_.cases | Where-Object { $_ }).Count -eq 0 })
        $publishedLibraries = @($definition.Libraries | Where-Object {
                (Resolve-BenchmarkLibrary -Policy $libraryPolicy -Name $_).restriction -eq 'open'
            })
        $publishableByPolicy = $inseparableLibraries.Count -eq 0
        $licenseRestriction = if ($inseparableLibraries.Count -gt 0) {
            (@($inseparableLibraries | ForEach-Object { "$($_.name): $($_.reason)" }) -join ' ')
        } else { $null }
        $leftOutLibraries = @($libraryPolicy.libraries | Where-Object {
                ($_.restriction -ne 'open' -and $_.id -notin $librarySelection.Enabled) -or $_.id -in $librarySelection.Excluded
            })
        $enabledLibraries = @($libraryPolicy.libraries | Where-Object { $_.id -in $librarySelection.Enabled })
        $expectedCases = @($definition.ExpectedCases | Where-Object {
                $case = $_
                @($leftOutLibraries | Where-Object { Test-BenchmarkCaseMatchesLibrary -Library $_ -Case $case }).Count -eq 0
            })
        # Values the libraries in this run read from the caller stay; values belonging only to libraries left out
        # of the run are hidden from it, and the plan shows them as cleared.
        $keptVariables = @(
            $workloadLibraries + $enabledLibraries | Where-Object { $_.restriction -eq 'open' -or $_.id -in $librarySelection.Enabled } |
                ForEach-Object { @($_.environmentVariables) + @($_.acceptance.environmentVariables) } | Where-Object { $_ }
        )
        $leftOutVariables = @($leftOutLibraries | ForEach-Object { @($_.acceptance.environmentVariables) + @($_.environmentVariables) } |
            Where-Object { $_ -and $_ -notin $keptVariables } | Select-Object -Unique)
        $buildProperties = [ordered]@{}
        foreach ($library in @($workloadLibraries | Where-Object { $_.id -in $librarySelection.Enabled -and $null -ne $_.acceptance })) {
            foreach ($property in @($library.acceptance.msbuildProperties.PSObject.Properties | Where-Object { $_ })) {
                $buildProperties[$property.Name] = [string] $property.Value
            }
        }
        # Framework, processor affinity and measurement profile are lane dimensions of one comparison, never part of its id.
        $laneParameters = [ordered]@{}
        if ($null -ne $affinityLabel) { $laneParameters.affinity = $affinityLabel.ToLowerInvariant() }
        if ($MeasurementProfile -ne 'Default') { $laneParameters.measurement = $MeasurementProfile.ToLowerInvariant() }
        $declaredComparison = @($declarations.comparisons | Where-Object comparisonId -eq $definition.ComparisonId)
        if ($declaredComparison.Count -ne 1) {
            throw "Workload '$name' ($($definition.ComparisonId)) is not in Build/benchmark-declarations.json; run ./Build/Export-BenchmarkDeclarations.ps1."
        }
        $declaredComparison = $declaredComparison[0]
        # Single-library comparisons and declared exclusions are recorded but never published.
        $declaredPublishable = @($declaredComparison.libraries).Count -gt 1 -and
            ($null -eq $declaredComparison.eligibility -or $declaredComparison.eligibility.publishable)
        $blockedReasons = @(
            if ($name -in @('markdownparse', 'markdownhtml') -and
                -not [string]::IsNullOrWhiteSpace($env:OFFICEIMO_MARKDOWN_BENCHMARK_CORPUS)) {
                'Clear OFFICEIMO_MARKDOWN_BENCHMARK_CORPUS to collect the declared seven-corpus CommonMark matrix.'
            }
            if ($Framework -notin @($declaredComparison.frameworks)) {
                "This comparison is declared for $(@($declaredComparison.frameworks) -join ', ') only: on $Framework a library runs another version or other cases."
            }
            foreach ($library in $workloadLibraries) {
                if ($library.restriction -eq 'excluded') {
                    "$($library.name) is excluded by the library policy: $($library.reason)"
                } elseif ($library.id -in $librarySelection.Excluded) {
                    "This comparison needs $($library.name), which -ExcludeLibrary removed."
                } elseif ($library.restriction -ne 'open' -and $library.id -notin $librarySelection.Enabled -and
                    @($library.cases).Count -eq 0) {
                    "$($library.name) is $($library.restriction): $($library.reason)"
                } elseif ($library.id -in $librarySelection.Enabled) {
                    Get-BenchmarkLibraryAcceptanceIssues -Library $library
                }
            }
            if ($Publish -and -not $publishableByPolicy -and $definition.CatalogEligible) {
                $licenseRestriction
            }
            if ($name -eq 'pdfcorpusread' -and [string]::IsNullOrWhiteSpace($PdfCorpusRoot)) {
                'PdfCorpusRoot is required; prepare the independent PDF corpus first.'
            }
            if ($AffinityMask -ne 0 -and $Platform -eq 'macos') { 'Nonzero affinity is unsupported on macOS.' }
        )
        $catalogEligibleByPolicy = -not $localOnlyRun -and $definition.CatalogEligible -and $publishableByPolicy
        $willCatalog = -not $localOnlyRun -and
            ($catalogEligibleByPolicy -or $definition.CollectNonPublishedEvidence) -and
            $blockedReasons.Count -eq 0
        $artifactsPath = Join-Path $OutputRoot "$Platform-$name-$BenchmarkProfile-$stamp"
        $commandOptions = @{
            Definition = $definition
            RepositoryRoot = $repositoryRoot
            Framework = $Framework
            BenchmarkProfile = $BenchmarkProfile
            ArtifactsPath = $artifactsPath
            BuildArtifactsPath = $BuildArtifactsPath
            KeepBenchmarkFiles = $KeepBenchmarkFiles
            BuildProperties = $buildProperties
            QuestPdfVersion = if ($name -in $questPdfWorkloads) { $effectiveQuestPdfVersion } else { $null }
        }
        $commands = @(Get-LibraryComparisonCommands @commandOptions)
        foreach ($command in $commands) {
            $quotedArguments = @($command.Arguments | ForEach-Object { "'" + ([string] $_).Replace("'", "''") + "'" })
            $command | Add-Member -NotePropertyName CommandLine -NotePropertyValue (
                $command.FilePath + ' ' + ($quotedArguments -join ' '))
        }
        $planItem = [pscustomobject]@{
            Workload = $name
            Area = $definition.Area
            Title = $definition.Title
            # Libraries that may appear in published results; restricted ones are listed separately.
            Libraries = $publishedLibraries
            DeclaredLibraries = @($definition.Libraries)
            ExpectedPlatforms = $definition.ExpectedPlatforms
            PublishRestriction = @(@($definition.PublishRestriction, $licenseRestriction) | Where-Object { $_ -and $_ -ne 'None' }) -join ' '
            RestrictedLibraries = @($restrictedLibraries.id)
            EnabledLibraries = @($librarySelection.Enabled)
            Platform = $Platform
            BlockedReasons = $blockedReasons
            CanRun = $blockedReasons.Count -eq 0
            Optional = $includeAll -and $name -notin $explicitWorkloads
            Commands = $commands
            CommandLines = @($commands.CommandLine)
            Environment = [ordered]@{
                OFFICEIMO_EXPECTED_BENCHMARK_AFFINITY = $affinityLabel
                OFFICEIMO_PDF_CORPUS_ROOT = if ($name -eq 'pdfcorpusread') { $PdfCorpusRoot } else { $null }
                HTMLTINKERX_PROJECT_PATH = if ($name -in @('pdfhtml', 'pdfhtmlpayload')) { $htmlTinkerXProjectPath } else { $null }
                OFFICEIMO_BENCHMARK_ENABLED_LIBRARIES = @($librarySelection.Enabled) -join ','
            }
            LeftOutVariables = $leftOutVariables
            # Enabled restricted libraries may add cases beyond ExpectedCases; the plan names them.
            OptionalCaseLibraries = @($enabledLibraries.id)
            ArtifactsPath = $artifactsPath
            BundleRoot = $staticRoot
            EvidenceCatalog = if ($willCatalog) { $catalogPath } else { $null }
            Filter = $definition.Filter
            ComparisonId = $definition.ComparisonId
            Framework = $Framework
            DeclaredFrameworks = @($declaredComparison.frameworks)
            Profile = $BenchmarkProfile
            LaneParameters = $laneParameters
            ProvenanceWorkloadId = $definition.ComparisonId
            CatalogEligible = $catalogEligibleByPolicy
            WillCatalog = $willCatalog
            Publish = [bool] $Publish -and $catalogEligibleByPolicy -and $declaredPublishable -and $willCatalog
            ExpectedCases = $expectedCases
            ExpectedCaseCount = $expectedCases.Count
            AffinityMask = $affinityLabel
            AffinityApplication = $effectiveAffinityApplication
            QuestPdfPackageVersion = if ($name -in $questPdfWorkloads) { $effectiveQuestPdfVersion } else { $null }
            InternalOnly = [bool] $localOnlyRun
        }
        foreach ($variable in $leftOutVariables) {
            $planItem.Environment[$variable] = $null
        }
        if ($MeasurementProfile -eq 'WarmedV1') {
            $planItem | Add-Member -NotePropertyName MeasurementProfile -NotePropertyValue 'WarmedV1'
            $planItem | Add-Member -NotePropertyName BenchmarkJobArguments -NotePropertyValue $definition.FullRunArguments
        }
        $planItem
    }
)
$executionPlanByWorkload = @{}
foreach ($item in $executionPlan) {
    $executionPlanByWorkload[$item.Workload] = $item
}

if ($isPlanning) {
    $executionPlan
    return
}

$outputPaths = @($staticRoot, $OutputRoot)
if (-not [string]::IsNullOrWhiteSpace($BuildArtifactsPath)) { $outputPaths += $BuildArtifactsPath }
Assert-BenchmarkOutputIgnored -RepositoryRoot $repositoryRoot -Path $outputPaths

$blockedRequired = @($executionPlan | Where-Object { -not $_.CanRun -and -not $_.Optional })
if ($blockedRequired.Count -gt 0) {
    throw (($blockedRequired | ForEach-Object { "$($_.Workload): $($_.BlockedReasons -join '; ')" }) -join [Environment]::NewLine)
}
foreach ($blocked in @($executionPlan | Where-Object { -not $_.CanRun -and $_.Optional })) {
    Write-Warning "Skipping optional $($blocked.Workload): $($blocked.BlockedReasons -join '; ')"
}
$selected = @($executionPlan | Where-Object CanRun | ForEach-Object Workload)
if ($selected.Count -eq 0) { throw 'No runnable workloads remain in the selection.' }

if ($usesHtmlTinkerX -and -not [string]::IsNullOrWhiteSpace($HtmlTinkerXRoot)) {
    if (-not (Test-Path -LiteralPath $htmlTinkerXProjectPath -PathType Leaf)) {
        throw "The HtmlTinkerX project was not found at '$htmlTinkerXProjectPath'."
    }

    $htmlTinkerXSourceRoot = (& git -C (Split-Path -Parent $htmlTinkerXProjectPath) rev-parse --show-toplevel).Trim()
    if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($htmlTinkerXSourceRoot)) {
        throw "Unable to resolve the HtmlTinkerX source root for '$htmlTinkerXProjectPath'."
    }
    $htmlTinkerXSourceCommit = (& git -C $htmlTinkerXSourceRoot rev-parse HEAD).Trim()
    if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($htmlTinkerXSourceCommit)) {
        throw 'Unable to resolve the HtmlTinkerX source commit for benchmark provenance.'
    }
    $htmlTinkerXSourceDirty = @(& git -C $htmlTinkerXSourceRoot status --porcelain --untracked-files=normal).Count -gt 0
}

function Get-LinuxThreadAffinityMasks {
    param([string] $TasksetPath, [int] $ProcessId)

    $output = @(& $TasksetPath -ap $ProcessId 2>&1)
    if ($LASTEXITCODE -ne 0 -or $output.Count -eq 0) {
        throw "Unable to read Linux thread affinity: $($output -join '; ')"
    }
    foreach ($line in $output) {
        $match = [regex]::Match([string] $line, '([0-9a-fA-F]+)\s*$')
        if (-not $match.Success) {
            throw "Unable to parse Linux thread affinity: $line"
        }
        [Convert]::ToUInt64($match.Groups[1].Value, 16)
    }
}

$benchmarkHost = $null
$originalProcessorAffinity = $null
$originalLinuxAffinity = $null
$linuxTaskset = $null
$previousNodeReuse = $env:MSBUILDDISABLENODEREUSE
$previousBuildServer = $env:DOTNET_CLI_USE_MSBUILD_SERVER
try {
$env:MSBUILDDISABLENODEREUSE = '1'
$env:DOTNET_CLI_USE_MSBUILD_SERVER = '0'

if ($AffinityMask -ne 0 -and $canInheritAffinity) {
    if ($isLinuxBenchmarkPlatform) {
        # Linux affinity is per thread. Process.ProcessorAffinity only pins the
        # calling thread, so a child launched by another PowerShell thread can
        # inherit the old mask. Pin every current thread before starting work.
        $linuxTaskset = (Get-Command taskset -CommandType Application -ErrorAction Stop | Select-Object -First 1).Source
        $initialMasks = @(Get-LinuxThreadAffinityMasks -TasksetPath $linuxTaskset -ProcessId $PID)
        if (@($initialMasks | Where-Object { $_ -ne $initialMasks[0] }).Count -ne 0) {
            throw 'The benchmark runner requires a uniform initial Linux thread affinity.'
        }
        $originalLinuxAffinity = '0x{0:X}' -f $initialMasks[0]
        $placementOutput = @(& $linuxTaskset -ap $affinityLabel $PID 2>&1)
        if ($LASTEXITCODE -ne 0) {
            throw "Unable to apply Linux affinity $affinityLabel`: $($placementOutput -join '; ')"
        }
        $appliedMasks = @(Get-LinuxThreadAffinityMasks -TasksetPath $linuxTaskset -ProcessId $PID)
        if (@($appliedMasks | Where-Object { $_ -ne $AffinityMask }).Count -ne 0) {
            throw "Requested Linux affinity $affinityLabel was not applied to every thread."
        }
    } else {
        $benchmarkHost = [System.Diagnostics.Process]::GetCurrentProcess()
        $originalProcessorAffinity = $benchmarkHost.ProcessorAffinity
        $benchmarkHost.ProcessorAffinity = [IntPtr]([long] $AffinityMask)
        $observedAffinity = [UInt64] $benchmarkHost.ProcessorAffinity.ToInt64()
        if ($observedAffinity -ne $AffinityMask) {
            throw "Requested affinity mask $affinityLabel, but Windows applied $('0x{0:X}' -f $observedAffinity)."
        }
    }
}

$catalogEligible = @($executionPlan | Where-Object WillCatalog).Count -gt 0
New-Item -ItemType Directory -Force -Path $OutputRoot | Out-Null
if ($catalogEligible) {
    New-Item -ItemType Directory -Force -Path $staticRoot | Out-Null
}

$gitSha = (& git -C $repositoryRoot rev-parse HEAD).Trim()
if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($gitSha)) {
    throw 'Unable to resolve the OfficeIMO source commit for benchmark provenance.'
}
$gitDirty = @(& git -C $repositoryRoot status --porcelain --untracked-files=normal).Count -gt 0
if ($catalogEligible -and $gitDirty) {
    throw 'Cataloged benchmark evidence requires a clean Git worktree so the recorded source commit identifies the measured code exactly.'
}
if ($selected -contains 'pdfhtml' -and $catalogEligible -and $htmlTinkerXSourceDirty) {
    throw 'Cataloged browser benchmark evidence requires a clean HtmlTinkerX worktree so its recorded source commit identifies the measured code exactly.'
}

& (Join-Path $PSScriptRoot 'Initialize-BenchmarkPowerForge.ps1') -PowerForgeRoot $PowerForgeRoot | Out-Null

$declaredLibraries = @(Get-BenchmarkDeclaredLibraries -Policy $libraryPolicy)
$subjectLibrary = [string] (@($declarations.comparisons)[0].subject)
$acceptedLibraries = @(
    foreach ($library in @($libraryPolicy.libraries | Where-Object { $_.id -in $librarySelection.Enabled })) {
        @($declaredLibraries | Where-Object PolicyId -eq $library.id | ForEach-Object Id)
    }
)
$outputs = [System.Collections.Generic.List[object]]::new()
$failures = [System.Collections.Generic.List[string]]::new()
foreach ($name in $selected) {
    $definition = $definitions[$name]
    $workloadPlan = $executionPlanByWorkload[$name]
    $artifactsPath = $workloadPlan.ArtifactsPath
    New-Item -ItemType Directory -Force -Path $artifactsPath | Out-Null
    $comparisonLibraries = @($definition.Libraries | ForEach-Object { Resolve-BenchmarkDeclaredLibrary -Libraries $declaredLibraries -Name $_ })

    # Restore first, so the competitor package versions and hashes the publish policy checks are bound into the
    # provenance before measuring.
    $projectPath = Join-Path $repositoryRoot $definition.Project
    $restoreArguments = @('restore', $projectPath, '--nologo', '--verbosity', 'quiet')
    if (-not [string]::IsNullOrWhiteSpace($BuildArtifactsPath)) {
        $restoreArguments += @('--artifacts-path', [System.IO.Path]::GetFullPath($BuildArtifactsPath))
    }
    $restoreArguments += @(@($workloadPlan.Commands | Where-Object Kind -eq 'measure')[0].Arguments | Where-Object { $_ -like '-p:*' })
    # The PDF comparisons build HtmlTinkerX from source or from its package depending on HTMLTINKERX_PROJECT_PATH;
    # restore with the same choice as the measurement, so the recorded packages are the ones that are timed.
    $isPdfComparison = $definition.Project -eq 'OfficeIMO.Pdf.Benchmarks.Comparisons\OfficeIMO.Pdf.Benchmarks.Comparisons.csproj'
    $previousRestoreHtmlTinkerX = $env:HTMLTINKERX_PROJECT_PATH
    try {
        if ($isPdfComparison) { $env:HTMLTINKERX_PROJECT_PATH = $htmlTinkerXProjectPath }
        & dotnet @restoreArguments | Out-Host
        $restoreExitCode = $LASTEXITCODE
    } finally {
        $env:HTMLTINKERX_PROJECT_PATH = $previousRestoreHtmlTinkerX
    }
    if ($restoreExitCode -ne 0) {
        $failures.Add("${name}: restoring $($definition.Project) failed with exit code $restoreExitCode.")
        continue
    }
    $projectName = [System.IO.Path]::GetFileNameWithoutExtension($projectPath)
    $assetsPath = if ([string]::IsNullOrWhiteSpace($BuildArtifactsPath)) {
        Join-Path (Split-Path -Parent $projectPath) 'obj/project.assets.json'
    } else {
        Join-Path ([System.IO.Path]::GetFullPath($BuildArtifactsPath)) "obj/$projectName/project.assets.json"
    }
    # Competitors and the packages their timed workflows depend on (such as HtmlAgilityPack for VersOne.Epub).
    $dependencyLibraries = @(@($definition.Dependencies | Where-Object { $_ }) | ForEach-Object { Resolve-BenchmarkDeclaredLibrary -Libraries $declaredLibraries -Name $_ })
    $packageMetadata = Get-BenchmarkPackageProvenance -AssetsPath $assetsPath -Framework $Framework `
        -Libraries @(@($comparisonLibraries | Where-Object Id -ne $subjectLibrary) + $dependencyLibraries)

    $provenanceMetadata = [ordered]@{
        'benchmark.workload.id' = $workloadPlan.ProvenanceWorkloadId
        'benchmark.workload.sourceCommit' = $gitSha
        'benchmark.workload.framework' = $Framework
        'benchmark.workload.requestedProcessPriority' = if ([string]::IsNullOrWhiteSpace($env:OFFICEIMO_BENCHMARK_PROCESS_PRIORITY)) { 'os-default' } else { $env:OFFICEIMO_BENCHMARK_PROCESS_PRIORITY }
    }
    if ($MeasurementProfile -eq 'WarmedV1') {
        $provenanceMetadata['benchmark.workload.measurementProfile'] = 'WarmedV1'
        $provenanceMetadata['benchmark.workload.jobArguments'] = ConvertTo-Json -InputObject $definition.FullRunArguments -Compress
    }
    if ($KeepBenchmarkFiles) {
        $provenanceMetadata['benchmark.workload.keepBenchmarkFiles'] = 'true'
        $provenanceMetadata['benchmark.workload.generateBinLog'] = 'true'
    }
    if ($definition.ValidatedSizeEvidence) {
        $provenanceMetadata['benchmark.workload.sizeEvidencePath'] = 'validated-size-evidence.json'
    }
    foreach ($package in $packageMetadata.GetEnumerator()) {
        $provenanceMetadata[$package.Key] = $package.Value
    }
    if ($name -in $questPdfWorkloads) {
        $provenanceMetadata['benchmark.dependency.questpdf.version'] = $effectiveQuestPdfVersion
        $provenanceMetadata['benchmark.dependency.questpdf.licenseType'] =
            if ([string]::IsNullOrWhiteSpace($env:OFFICEIMO_QUESTPDF_LICENSE_TYPE)) { 'Community' } else { $env:OFFICEIMO_QUESTPDF_LICENSE_TYPE }
        if (-not $alternateQuestPdf) {
            $provenanceMetadata['benchmark.dependency.questpdf.license'] = 'MIT'
            $provenanceMetadata['benchmark.dependency.questpdf.licenseUrl'] =
                'https://github.com/QuestPDF/QuestPDF/blob/2026.5.0/LICENSE.md'
        } else {
            $provenanceMetadata['benchmark.dependency.questpdf.scope'] = 'local-only; licence accepted by the caller'
        }
    }
    $provenanceMetadata['benchmark.workload.affinityApplication'] = $effectiveAffinityApplication
    if ($null -ne $affinityLabel) {
        $provenanceMetadata['benchmark.workload.affinityMask'] = $affinityLabel
    }
    if ($name -in @('pdfhtml', 'pdfhtmlpayload')) {
        [xml] $pdfComparisonProject = Get-Content -LiteralPath (
            Join-Path $repositoryRoot 'OfficeIMO.Pdf.Benchmarks.Comparisons/OfficeIMO.Pdf.Benchmarks.Comparisons.csproj') -Raw
        $peachVersionNode = $pdfComparisonProject.SelectSingleNode('//PeachPdfBenchmarkVersion')
        if ($null -eq $peachVersionNode -or [string] $peachVersionNode.InnerText -notmatch '^\d+\.\d+\.\d+$') {
            throw 'The HTML comparison must declare an exact PeachPDF package version.'
        }
        $provenanceMetadata['benchmark.fixture.html.font.profile'] = 'SharedCarlitoRegularBold'
        $peachVersion = [string] $peachVersionNode.InnerText
        $provenanceMetadata['benchmark.dependency.peachpdf.version'] = $peachVersion
        $pdfPeckerVersionNode = $pdfComparisonProject.SelectSingleNode('//PdfPeckerBenchmarkVersion')
        if ($null -eq $pdfPeckerVersionNode -or [string] $pdfPeckerVersionNode.InnerText -notmatch '^\d+\.\d+\.\d+$') {
            throw 'The HTML comparison must declare an exact PdfPecker package version.'
        }
        $pdfPeckerVersion = [string] $pdfPeckerVersionNode.InnerText
        $provenanceMetadata['benchmark.dependency.pdfpecker.version'] = $pdfPeckerVersion
        $provenanceMetadata['benchmark.dependency.pdfpecker.license'] = 'MIT AND OFL-1.1 AND Bitstream-Vera AND APAFML AND BSD-3-Clause AND MIT-Modern-Variant AND Unicode-3.0 AND FSFAP AND LPPL-1.3c'
        $provenanceMetadata['benchmark.dependency.pdfpecker.licenseUrl'] =
            "https://www.nuget.org/packages/PdfPecker/$pdfPeckerVersion/License"
        # The shared capture reserves an empty raw directory. Probe the package
        # beside it, then admit the probe as a captured artifact after reservation.
        $packageEvidencePath = Join-Path $OutputRoot (
            '{0}-pdfpecker-package.json' -f (Split-Path -Leaf $artifactsPath))
        $packageEvidenceArguments = @($workloadPlan.Commands | Where-Object Kind -eq 'package-evidence')[0].Arguments
        & dotnet @packageEvidenceArguments
        if ($LASTEXITCODE -ne 0) { throw 'Could not inspect the restored PdfPecker comparison package.' }
        $packageEvidence = Get-Content -LiteralPath $packageEvidencePath -Raw | ConvertFrom-Json
        if ($packageEvidence.id -cne 'PdfPecker' -or $packageEvidence.version -cne $pdfPeckerVersion) {
            throw 'The restored PdfPecker comparison package does not match the declared version.'
        }
        $provenanceMetadata['benchmark.dependency.pdfpecker.license'] = [string] $packageEvidence.license
        foreach ($field in @('sourceCommit', 'licenseSha256', 'thirdPartyNoticesSha256', 'packageSha512')) {
            $provenanceMetadata["benchmark.dependency.pdfpecker.$field"] = [string] $packageEvidence.$field
        }
        foreach ($face in @('Regular', 'Bold')) {
            $fontPath = Join-Path $repositoryRoot "Website/Apps/OfficeIMO.Web.Converter/Assets/Fonts/Carlito-$face.ttf"
            $provenanceMetadata["benchmark.fixture.html.font.$($face.ToLowerInvariant()).sha256"] =
                (Get-FileHash -LiteralPath $fontPath -Algorithm SHA256).Hash.ToLowerInvariant()
        }
    }
    if ($name -eq 'pdfhtml') {
        if ($null -ne $htmlTinkerXSourceCommit) {
            $provenanceMetadata['benchmark.browser.owner'] = 'HtmlTinkerX source'
            $provenanceMetadata['benchmark.browser.sourceCommit'] = $htmlTinkerXSourceCommit
        } else {
            $htmlTinkerXReferences = @($pdfComparisonProject.SelectNodes('//PackageReference[@Include="HtmlTinkerX"]'))
            if ($htmlTinkerXReferences.Count -ne 1 -or
                [string] $htmlTinkerXReferences[0].Version -notmatch '^\d+\.\d+\.\d+$') {
                throw 'The HTML comparison must declare one exact HtmlTinkerX package version.'
            }
            $htmlTinkerXPackageVersion = [string] $htmlTinkerXReferences[0].Version
            $provenanceMetadata['benchmark.browser.owner'] = 'HtmlTinkerX package'
            $provenanceMetadata['benchmark.browser.packageVersion'] = $htmlTinkerXPackageVersion
            $provenanceMetadata['benchmark.dependency.htmltinkerx.version'] = $htmlTinkerXPackageVersion
        }
    }
    $provenanceCapture = Start-BenchmarkProvenanceCapture `
        -SourceRoot $repositoryRoot `
        -ArtifactRoot $artifactsPath `
        -Metadata $provenanceMetadata `
        -RunMode $BenchmarkProfile
    if ($name -in @('pdfhtml', 'pdfhtmlpayload')) {
        Move-Item -LiteralPath $packageEvidencePath -Destination (
            Join-Path $artifactsPath 'pdfpecker-package.json')
    }

    $arguments = @($workloadPlan.Commands | Where-Object Kind -eq 'measure')[0].Arguments

    Push-Location -LiteralPath $repositoryRoot
    $previousPdfCorpusRoot = $env:OFFICEIMO_PDF_CORPUS_ROOT
    $previousHtmlTinkerXProjectPath = $env:HTMLTINKERX_PROJECT_PATH
    $previousExpectedAffinity = $env:OFFICEIMO_EXPECTED_BENCHMARK_AFFINITY
    $previousEnabledLibraries = $env:OFFICEIMO_BENCHMARK_ENABLED_LIBRARIES
    $previousValidationPath = $env:OFFICEIMO_BENCHMARK_VALIDATION_PATH
    # Benchmarks append one JSON object per validated output ({ library, case, status, detail }) to this file
    # before timing; it stays inside the captured artifacts.
    $validationLog = Join-Path $artifactsPath 'validation.jsonl'
    $previousLeftOutValues = @{}
    foreach ($variable in $workloadPlan.LeftOutVariables) {
        $previousLeftOutValues[$variable] = [Environment]::GetEnvironmentVariable($variable)
    }
    try {
        # Benchmark code runs only the restricted libraries listed here. Licence values belonging to libraries
        # left out of this run are hidden from it; enabled libraries read the caller's own values.
        $env:OFFICEIMO_BENCHMARK_ENABLED_LIBRARIES = $workloadPlan.Environment.OFFICEIMO_BENCHMARK_ENABLED_LIBRARIES
        foreach ($variable in $workloadPlan.LeftOutVariables) {
            [Environment]::SetEnvironmentVariable($variable, $null)
        }
        if ($name -eq 'pdfcorpusread') {
            $env:OFFICEIMO_PDF_CORPUS_ROOT = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($PdfCorpusRoot)
        }
        if ($isPdfComparison) {
            $env:HTMLTINKERX_PROJECT_PATH = $htmlTinkerXProjectPath
        }
        $env:OFFICEIMO_EXPECTED_BENCHMARK_AFFINITY = $affinityLabel
        $env:OFFICEIMO_BENCHMARK_VALIDATION_PATH = $validationLog
        & dotnet @arguments
        $benchmarkExitCode = $LASTEXITCODE
    } finally {
        $env:OFFICEIMO_PDF_CORPUS_ROOT = $previousPdfCorpusRoot
        $env:HTMLTINKERX_PROJECT_PATH = $previousHtmlTinkerXProjectPath
        $env:OFFICEIMO_EXPECTED_BENCHMARK_AFFINITY = $previousExpectedAffinity
        $env:OFFICEIMO_BENCHMARK_ENABLED_LIBRARIES = $previousEnabledLibraries
        $env:OFFICEIMO_BENCHMARK_VALIDATION_PATH = $previousValidationPath
        foreach ($variable in $previousLeftOutValues.Keys) {
            [Environment]::SetEnvironmentVariable($variable, $previousLeftOutValues[$variable])
        }
        Pop-Location
    }
    if ($benchmarkExitCode -ne 0) {
        # A crashed or failed case is a result: its cells are recorded as failed and the rest of the run continues.
        $failures.Add("${name}: BenchmarkDotNet exited with code $benchmarkExitCode; failed cases are recorded as failed cells.")
    }

    $sizeEvidencePath = $null
    if ($definition.ValidatedSizeEvidence) {
        $sizeEvidencePath = Join-Path $artifactsPath 'validated-size-evidence.json'
        $sizeArguments = @($workloadPlan.Commands | Where-Object Kind -eq 'validate-size')[0].Arguments
        Push-Location -LiteralPath $repositoryRoot
        try {
            & dotnet @sizeArguments
            $sizeEvidenceExitCode = $LASTEXITCODE
        } finally {
            Pop-Location
        }
        if ($sizeEvidenceExitCode -ne 0) {
            throw "$name validated size evidence failed with exit code $sizeEvidenceExitCode."
        }

        $sizeResults = Get-Content -LiteralPath $sizeEvidencePath -Raw | ConvertFrom-Json
        $sizeEvidence = [ordered]@{
            Workload = $name
            ComparisonId = $definition.ComparisonId
            SourceCommit = $gitSha
            Framework = $Framework
            Platform = $platform
            CapturedAtUtc = [DateTimeOffset]::UtcNow.ToString('O')
            Results = $sizeResults
        }
        $sizeEvidenceJson = $sizeEvidence | ConvertTo-Json -Depth 12
        [System.IO.File]::WriteAllText(
            $sizeEvidencePath,
            $sizeEvidenceJson,
            [System.Text.UTF8Encoding]::new($false))
    }
    $provenanceCapture |
        Complete-BenchmarkProvenanceCapture |
        Out-Null

    # Derived files stay outside the completed raw-artifact provenance boundary.
    $importOptions = @{
        Path = $artifactsPath
        Suite = $definition.Suite
    }
    $importMaps = Get-BenchmarkComparisonImportMaps -Workload $name -Definition $definition -Libraries $declaredLibraries
    if ($importMaps.LibraryMap.Count -gt 0) { $importOptions.LibraryMap = $importMaps.LibraryMap }
    if ($importMaps.CaseMap.Count -gt 0) { $importOptions.CaseMap = $importMaps.CaseMap }
    if ($acceptedLibraries.Count -gt 0) { $importOptions.AcceptLicense = $acceptedLibraries }
    if (Test-Path -LiteralPath $validationLog -PathType Leaf) {
        $validations = @(Get-Content -LiteralPath $validationLog -Encoding UTF8 | Where-Object { $_.Trim() } | ForEach-Object { $_ | ConvertFrom-Json })
        $validationPath = Join-Path $OutputRoot ('{0}-validation.json' -f (Split-Path -Leaf $artifactsPath))
        ConvertTo-Json -InputObject $validations -Depth 5 | Set-Content -LiteralPath $validationPath -Encoding utf8NoBOM
        $importOptions.ValidationPath = $validationPath
    }
    try {
        $result = Import-BenchmarkResult @importOptions
    } catch {
        $failures.Add("${name}: the results could not be imported: $($_.Exception.Message)")
        continue
    }
    $normalizedPath = Join-Path $OutputRoot ('{0}-normalized-result.json' -f (Split-Path -Leaf $artifactsPath))
    Write-BenchmarkEvidenceResult -Path $normalizedPath -InputObject $result

    # Drift between the declarations and what ran blocks publication; the lane is still recorded. PowerForge also
    # reports every declared case, library and axis-value combination (or declared cell) that did not run, and every
    # combination that ran outside the declared cells.
    $drift = @(Test-BenchmarkDeclaration -Path $declarationPath -InputObject $result -ComparisonId $definition.ComparisonId -PassThru)
    $publishLane = $workloadPlan.Publish
    if ($drift.Count -gt 0) {
        $failures.Add("${name}: the results drifted from Build/benchmark-declarations.json ($(@($drift | ForEach-Object { $_.Message }) -join ' ')). Regenerate the declarations or fix the benchmark.")
        $publishLane = $false
    }

    $laneKey = $null
    if ($workloadPlan.WillCatalog) {
        $resultName = @($definition.ComparisonId, $Platform, $Framework, $BenchmarkProfile) + @($workloadPlan.LaneParameters.Values)
        $updateOptions = @{
            InputObject = $result
            Path = $catalogPath
            ComparisonId = $definition.ComparisonId
            ResultPath = ($resultName -join '-') + '.json'
            Profile = $BenchmarkProfile
            Framework = $Framework
            Platform = $Platform
            DeclarationPath = $declarationPath
        }
        if ($workloadPlan.LaneParameters.Count -gt 0) {
            $updateOptions.LaneParameter = [hashtable] $workloadPlan.LaneParameters
        }
        if (-not $portableBundle) {
            $updateOptions.ResultArtifactPath = Join-Path $staticRoot $updateOptions.ResultPath
            $updateOptions.ResultPath = '/data/benchmarks/evidence/' + $updateOptions.ResultPath
        }
        try {
            $catalog = Update-BenchmarkEvidenceCatalog @updateOptions -Publish:$publishLane
        } catch {
            if (-not $publishLane) { throw }
            # The lane does not meet the publish policy; its evidence is kept unpublished and the reason reported.
            $failures.Add("${name}: not published: $($_.Exception.Message)")
            $publishLane = $false
            $catalog = Update-BenchmarkEvidenceCatalog @updateOptions
        }
        $entry = @($catalog.Entries | Where-Object {
                $_.ComparisonId -eq $definition.ComparisonId -and $_.Platform -eq $Platform -and $_.Framework -eq $Framework
            } | Sort-Object GeneratedUtc -Descending | Select-Object -First 1)
        if ($entry.Count -eq 1) { $laneKey = $entry[0].LaneKey }
    }

    $outputs.Add([pscustomobject]@{
        Workload = $name
        Area = $definition.Area
        ComparisonId = $definition.ComparisonId
        Platform = $Platform
        Framework = $Framework
        Profile = $BenchmarkProfile
        LaneParameters = $workloadPlan.LaneParameters
        LaneKey = $laneKey
        Publish = $publishLane
        Drift = @($drift | ForEach-Object { $_.Message })
        AffinityMask = $affinityLabel
        AffinityApplication = $effectiveAffinityApplication
        QuestPdfPackageVersion = $workloadPlan.QuestPdfPackageVersion
        InternalOnly = $workloadPlan.InternalOnly
        SourceCommit = $gitSha
        ArtifactsPath = $artifactsPath
        NormalizedResult = $normalizedPath
        SizeEvidence = $sizeEvidencePath
        BundleRoot = $staticRoot
        EvidenceCatalog = if ($workloadPlan.WillCatalog) { $catalogPath } else { $null }
    })
}

$outputs
if ($failures.Count -gt 0) {
    throw ("Some workloads did not complete cleanly; the evidence that was measured is recorded:" + [Environment]::NewLine +
        ($failures -join [Environment]::NewLine))
}
} finally {
    $env:MSBUILDDISABLENODEREUSE = $previousNodeReuse
    $env:DOTNET_CLI_USE_MSBUILD_SERVER = $previousBuildServer
    if ($null -ne $linuxTaskset -and $null -ne $originalLinuxAffinity) {
        $restoreOutput = @(& $linuxTaskset -ap $originalLinuxAffinity $PID 2>&1)
        if ($LASTEXITCODE -ne 0) {
            throw "The benchmark runner could not restore Linux affinity: $($restoreOutput -join '; ')"
        }
        $restoredMasks = @(Get-LinuxThreadAffinityMasks -TasksetPath $linuxTaskset -ProcessId $PID)
        $expectedRestoreMask = [Convert]::ToUInt64($originalLinuxAffinity.Substring(2), 16)
        if (@($restoredMasks | Where-Object { $_ -ne $expectedRestoreMask }).Count -ne 0) {
            throw 'The benchmark runner could not restore every Linux thread affinity.'
        }
    } elseif ($null -ne $benchmarkHost -and $null -ne $originalProcessorAffinity) {
        $benchmarkHost.ProcessorAffinity = $originalProcessorAffinity
        if ($benchmarkHost.ProcessorAffinity.ToInt64() -ne $originalProcessorAffinity.ToInt64()) {
            throw 'The benchmark runner could not restore the invoking process affinity.'
        }
    }
}
