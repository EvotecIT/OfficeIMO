param(
    [ValidateSet('quick', 'full')]
    [string] $RunMode = 'full',
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
    [string] $QuestPdfPackageVersion,
    [ValidateSet('Community', 'Evaluation', 'Professional', 'Enterprise')]
    [string] $QuestPdfLicenseType = 'Community',
    [switch] $InternalQuestPdf,
    [switch] $ConfirmQuestPdfAuthorization,
    [switch] $AcceptNPOIOSMFLicense,
    [switch] $Publish,
    [switch] $PlanOnly,
    [switch] $KeepBenchmarkFiles,
    [ValidateSet('Default', 'WarmedV1')]
    [string] $MeasurementProfile = 'Default'
)

$ErrorActionPreference = 'Stop'

if ($Publish -and $RunMode -ne 'full') {
    throw 'Only a full BenchmarkDotNet run can be marked publishable.'
}

if ($MeasurementProfile -eq 'WarmedV1' -and
    ($RunMode -ne 'full' -or $Workload.Count -ne 1 -or $Workload[0] -ne 'pdfhtmlpayload')) {
    throw 'WarmedV1 requires -RunMode full -Workload pdfhtmlpayload.'
}

. (Join-Path $PSScriptRoot 'BenchmarkEvidence.ps1')

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
$websiteRoot = Join-Path $repositoryRoot 'Website/static/data/benchmarks/library-comparisons'
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
$comparisonPath = Join-Path $staticRoot 'comparisons.json'

. (Join-Path $PSScriptRoot 'LibraryComparisonDefinitions.ps1')
$definitions = Get-LibraryComparisonDefinitions -Framework $Framework

if ($MeasurementProfile -eq 'WarmedV1') {
    $definitions.pdfhtmlpayload.ComparisonId += '-warmed-v1'
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
        if ($requested -eq 'word') { 'wordcreate'; 'wordreport'; 'wordread'; 'wordreplace' }
        elseif ($requested -eq 'powerpoint') { 'pptxcreate'; 'pptxedit' }
        elseif ($definitions.Contains($requested)) { $requested }
        else { throw "Unknown workload '$requested'. Choose all, word, powerpoint, or: $($definitions.Keys -join ', ')." }
    }
)
$selected = @(@(if ($includeAll) { $definitions.Keys }; $explicitWorkloads) | Select-Object -Unique)
if ($selected.Count -eq 0) { throw 'Select at least one workload.' }
$containsQuestPdfWorkload = @($selected | Where-Object { $_ -in $questPdfWorkloads }).Count -gt 0
$effectiveQuestPdfVersion = if ($InternalQuestPdf) { $QuestPdfPackageVersion } else { $questPdfBenchmarkVersion }

if ($InternalQuestPdf) {
    if ($Publish) {
        throw 'Internal QuestPDF comparisons cannot be published or written to the website benchmark catalog.'
    }
    if (-not $containsQuestPdfWorkload -or @($selected | Where-Object { $_ -notin $questPdfWorkloads }).Count -gt 0) {
        throw 'Internal QuestPDF mode accepts only workloads that actually exercise QuestPDF.'
    }
    if ([string]::IsNullOrWhiteSpace($QuestPdfPackageVersion)) {
        throw 'Internal QuestPDF mode requires an explicit -QuestPdfPackageVersion.'
    }
} elseif (-not [string]::IsNullOrWhiteSpace($QuestPdfPackageVersion)) {
    throw '-QuestPdfPackageVersion is available only with -InternalQuestPdf.'
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
        $catalogComparisonId = if ($null -ne $affinityLabel) {
            '{0}-affinity-{1}' -f $definition.ComparisonId, $affinityLabel.ToLowerInvariant()
        } else {
            $definition.ComparisonId
        }
        $blockedReasons = @(
            if ($name -in @('markdownparse', 'markdownhtml') -and
                -not [string]::IsNullOrWhiteSpace($env:OFFICEIMO_MARKDOWN_BENCHMARK_CORPUS)) {
                'Clear OFFICEIMO_MARKDOWN_BENCHMARK_CORPUS to collect the declared seven-corpus CommonMark matrix.'
            }
            if ($null -ne $definition.Frameworks -and $Framework -notin $definition.Frameworks) {
                "This comparison requires $($definition.Frameworks -join ', '); ShapeCrawler 0.82.0 ships only net10.0."
            }
            if ($name -like 'word*') {
                if ($Publish) { $definition.PublishRestriction }
                if (-not $AcceptNPOIOSMFLicense) {
                    'NPOI 2.8.1 OSMF EULA acknowledgement required: -AcceptNPOIOSMFLicense (https://github.com/nissl-lab/npoi/blob/master/OSMFEULA.txt).'
                }
            }
            if ($InternalQuestPdf -and -not $ConfirmQuestPdfAuthorization) {
                'Internal QuestPDF requires -ConfirmQuestPdfAuthorization for the selected version.'
            }
            if ($name -eq 'pdfcorpusread' -and [string]::IsNullOrWhiteSpace($PdfCorpusRoot)) {
                'PdfCorpusRoot is required; prepare the independent PDF corpus first.'
            }
            if ($AffinityMask -ne 0 -and $Platform -eq 'macos') { 'Nonzero affinity is unsupported on macOS.' }
        )
        $catalogEligibleByPolicy = -not $InternalQuestPdf -and $definition.CatalogEligible
        $willCatalog = -not $InternalQuestPdf -and
            ($catalogEligibleByPolicy -or $definition.CollectNonPublishedEvidence) -and
            $blockedReasons.Count -eq 0
        $artifactsPath = Join-Path $OutputRoot "$Platform-$name-$RunMode-$stamp"
        $commandOptions = @{
            Definition = $definition
            RepositoryRoot = $repositoryRoot
            Framework = $Framework
            RunMode = $RunMode
            ArtifactsPath = $artifactsPath
            BuildArtifactsPath = $BuildArtifactsPath
            KeepBenchmarkFiles = $KeepBenchmarkFiles
            Word = $name -like 'word*'
            QuestPdfVersion = if ($name -in $questPdfWorkloads) { $effectiveQuestPdfVersion } else { $null }
            InternalQuestPdf = $InternalQuestPdf
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
            Libraries = $definition.Libraries
            ExpectedPlatforms = $definition.ExpectedPlatforms
            PublishRestriction = $definition.PublishRestriction
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
                OFFICEIMO_QUESTPDF_LICENSE_TYPE = if ($name -in $questPdfWorkloads) { $QuestPdfLicenseType } else { $null }
            }
            ArtifactsPath = $artifactsPath
            BundleRoot = $staticRoot
            ComparisonPath = if (-not $InternalQuestPdf) { $comparisonPath } else { $null }
            EvidenceCatalog = if ($willCatalog) { $catalogPath } else { $null }
            Filter = $definition.Filter
            ComparisonId = $definition.ComparisonId
            CatalogComparisonId = $catalogComparisonId
            ProvenanceWorkloadId = $catalogComparisonId
            CatalogEligible = $catalogEligibleByPolicy
            WillCatalog = $willCatalog
            Publish = [bool] $Publish -and $catalogEligibleByPolicy -and $willCatalog
            ExpectedCaseCount = @($definition.ExpectedCases).Count
            AffinityMask = $affinityLabel
            AffinityApplication = $effectiveAffinityApplication
            QuestPdfPackageVersion = if ($name -in $questPdfWorkloads) { $effectiveQuestPdfVersion } else { $null }
            InternalOnly = [bool] $InternalQuestPdf
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

$measurements = [System.Collections.Generic.List[object]]::new()
foreach ($name in $selected) {
    $definition = $definitions[$name]
    $workloadPlan = $executionPlanByWorkload[$name]
    $artifactsPath = $workloadPlan.ArtifactsPath
    New-Item -ItemType Directory -Force -Path $artifactsPath | Out-Null
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
    if ($name -in $questPdfWorkloads) {
        $provenanceMetadata['benchmark.dependency.questpdf.version'] = $effectiveQuestPdfVersion
        $provenanceMetadata['benchmark.dependency.questpdf.licenseType'] = $QuestPdfLicenseType
        if (-not $InternalQuestPdf) {
            $provenanceMetadata['benchmark.dependency.questpdf.license'] = 'MIT'
            $provenanceMetadata['benchmark.dependency.questpdf.licenseUrl'] =
                'https://github.com/QuestPDF/QuestPDF/blob/2026.5.0/LICENSE.md'
        } else {
            $provenanceMetadata['benchmark.dependency.questpdf.scope'] = 'internal-only; maintainer-confirmed authorization'
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
        $provenanceMetadata['benchmark.package.peachpdf.version'] = $peachVersion
        $provenanceMetadata['benchmark.dependency.peachpdf.version'] = $peachVersion
        $pdfPeckerVersionNode = $pdfComparisonProject.SelectSingleNode('//PdfPeckerBenchmarkVersion')
        if ($null -eq $pdfPeckerVersionNode -or [string] $pdfPeckerVersionNode.InnerText -notmatch '^\d+\.\d+\.\d+$') {
            throw 'The HTML comparison must declare an exact PdfPecker package version.'
        }
        $pdfPeckerVersion = [string] $pdfPeckerVersionNode.InnerText
        $provenanceMetadata['benchmark.package.pdfpecker.version'] = $pdfPeckerVersion
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
            $provenanceMetadata['benchmark.package.htmltinkerx.version'] = $htmlTinkerXPackageVersion
            $provenanceMetadata['benchmark.dependency.htmltinkerx.version'] = $htmlTinkerXPackageVersion
        }
    }
    $provenanceCapture = Start-BenchmarkProvenanceCapture `
        -SourceRoot $repositoryRoot `
        -ArtifactRoot $artifactsPath `
        -Metadata $provenanceMetadata `
        -RunMode $RunMode
    if ($name -in @('pdfhtml', 'pdfhtmlpayload')) {
        Move-Item -LiteralPath $packageEvidencePath -Destination (
            Join-Path $artifactsPath 'pdfpecker-package.json')
    }

    $arguments = @($workloadPlan.Commands | Where-Object Kind -eq 'measure')[0].Arguments
    $isPdfComparison = $definition.Project -eq 'OfficeIMO.Pdf.Benchmarks.Comparisons\OfficeIMO.Pdf.Benchmarks.Comparisons.csproj'

    Push-Location -LiteralPath $repositoryRoot
    $previousPdfCorpusRoot = $env:OFFICEIMO_PDF_CORPUS_ROOT
    $previousHtmlTinkerXProjectPath = $env:HTMLTINKERX_PROJECT_PATH
    $previousExpectedAffinity = $env:OFFICEIMO_EXPECTED_BENCHMARK_AFFINITY
    $previousQuestPdfLicenseType = $env:OFFICEIMO_QUESTPDF_LICENSE_TYPE
    try {
        if ($name -eq 'pdfcorpusread') {
            $env:OFFICEIMO_PDF_CORPUS_ROOT = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($PdfCorpusRoot)
        }
        if ($isPdfComparison) {
            $env:HTMLTINKERX_PROJECT_PATH = $htmlTinkerXProjectPath
        }
        if ($name -in $questPdfWorkloads) {
            $env:OFFICEIMO_QUESTPDF_LICENSE_TYPE = $QuestPdfLicenseType
        }
        $env:OFFICEIMO_EXPECTED_BENCHMARK_AFFINITY = $affinityLabel
        & dotnet @arguments
        $benchmarkExitCode = $LASTEXITCODE
    } finally {
        $env:OFFICEIMO_PDF_CORPUS_ROOT = $previousPdfCorpusRoot
        $env:HTMLTINKERX_PROJECT_PATH = $previousHtmlTinkerXProjectPath
        $env:OFFICEIMO_EXPECTED_BENCHMARK_AFFINITY = $previousExpectedAffinity
        $env:OFFICEIMO_QUESTPDF_LICENSE_TYPE = $previousQuestPdfLicenseType
        Pop-Location
    }
    if ($benchmarkExitCode -ne 0) {
        throw "$name benchmark run failed with exit code $benchmarkExitCode."
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

    $result = Import-BenchmarkResult -Path $artifactsPath -Suite $definition.Suite
    $successfulCases = @(
        $result.Summary |
            Where-Object {
                $_.Status -in @('Success', 'Succeeded') -and
                $_.SampleCount -gt 0 -and
                $null -ne $_.MedianMs
            } |
            ForEach-Object {
                Get-BenchmarkEvidenceCaseIdentity `
                    -Row $_ `
                    -VariableName $definition.IdentityVariables
            } |
            Sort-Object -Unique
    )
    $missingCases = @(
        $definition.ExpectedCases |
            Where-Object { $_ -notin $successfulCases }
    )
    $unexpectedCases = @(
        $successfulCases |
            Where-Object { $_ -notin $definition.ExpectedCases }
    )
    if ($missingCases.Count -gt 0 -or $unexpectedCases.Count -gt 0) {
        $details = @()
        if ($missingCases.Count -gt 0) {
            $details += "missing successful cases: $($missingCases -join ', ')"
        }
        if ($unexpectedCases.Count -gt 0) {
            $details += "unexpected cases: $($unexpectedCases -join ', ')"
        }
        throw "$name benchmark evidence is incomplete ($($details -join '; '))."
    }

    $evidenceLocation = Get-BenchmarkEvidenceLocation `
        -ComparisonId $workloadPlan.CatalogComparisonId `
        -Platform $platform `
        -RunMode $RunMode `
        -StaticRoot $staticRoot `
        -Portable:$portableBundle
    # Keep derived output outside the completed raw-artifact provenance boundary.
    $normalizedPath = Join-Path $OutputRoot (
        '{0}-normalized-result.json' -f (Split-Path -Leaf $artifactsPath))
    Write-BenchmarkEvidenceResult -Path $normalizedPath -InputObject $result

    $measurements.Add([pscustomobject]@{
        Workload = $name
        Definition = $definition
        Result = $result
        EvidenceLocation = $evidenceLocation
        ArtifactsPath = $artifactsPath
        NormalizedResult = $normalizedPath
        SizeEvidence = $sizeEvidencePath
        CatalogEligible = $workloadPlan.WillCatalog
    })
}

if (($measurements.Count -ne $selected.Count) -or
    (@($measurements | Where-Object { $null -eq $_.Result }).Count -gt 0)) {
    throw 'Benchmark measurement collection did not produce exactly one normalized result per selected workload.'
}

if ($catalogEligible) {
    foreach ($measurement in @($measurements | Where-Object CatalogEligible)) {
        Update-BenchmarkEvidenceCatalog `
            -InputObject $measurement.Result `
            -Path $catalogPath `
            -ComparisonId $executionPlanByWorkload[$measurement.Workload].CatalogComparisonId `
            -ResultPath $measurement.EvidenceLocation.ResultPath `
            -ResultArtifactPath $measurement.EvidenceLocation.Path `
            -RunMode $RunMode `
            -Platform $platform `
            -ExpectedPlatform $measurement.Definition.ExpectedPlatforms `
            -Publish:($executionPlanByWorkload[$measurement.Workload].Publish) | Out-Null
    }
}

if (-not $InternalQuestPdf) {
    # Declarations include restricted and uncollected workloads, even for a subset run.
    # Keep this metadata outside completed raw provenance; PowerForge owns catalog merging.
    $catalogDeclarations = [ordered]@{}
    if (Test-Path -LiteralPath $catalogPath -PathType Leaf) {
        $existingCatalog = Get-Content -LiteralPath $catalogPath -Raw | ConvertFrom-Json
        foreach ($declaration in @($existingCatalog.comparisons)) {
            $catalogDeclarations[$declaration.comparisonId] = $declaration
        }
    }
    $baseDefinitions = Get-LibraryComparisonDefinitions -Framework $Framework
    foreach ($definition in $baseDefinitions.Values) {
        if ($definition.Frameworks -and $Framework -notin $definition.Frameworks) { continue }
        $catalogDeclarations[$definition.ComparisonId] = [pscustomobject]@{
            comparisonId = $definition.ComparisonId
            area = $definition.Area
            title = $definition.Title
            libraries = @($definition.Libraries)
            expectedPlatforms = @($definition.ExpectedPlatforms)
        }
    }
    foreach ($item in $executionPlan) {
        $definition = $baseDefinitions[$item.Workload]
        if ($definition.Frameworks -and $Framework -notin $definition.Frameworks) { continue }
        $baseId = $definition.ComparisonId
        if ($item.CatalogComparisonId -eq $baseId) { continue }
        $catalogDeclarations[$item.CatalogComparisonId] = [pscustomobject]@{
            comparisonId = $item.CatalogComparisonId
            area = $item.Area
            title = $item.Title + ' (' + $item.CatalogComparisonId.Substring($baseId.Length).TrimStart('-') + ')'
            libraries = @($item.Libraries)
            expectedPlatforms = @($item.ExpectedPlatforms)
        }
    }
    New-Item -ItemType Directory -Force -Path $staticRoot | Out-Null
    ConvertTo-Json -InputObject @($catalogDeclarations.Values) -Depth 10 |
        Set-Content -LiteralPath $comparisonPath -Encoding utf8NoBOM
    if (-not (Test-Path -LiteralPath $catalogPath -PathType Leaf)) {
        Write-BenchmarkEvidenceResult -Path $catalogPath -InputObject ([PowerForge.BenchmarkEvidenceCatalog]::new())
    }
    $mergedCatalog = Merge-BenchmarkEvidenceCatalog -SourcePath $catalogPath -Path $catalogPath -ComparisonPath $comparisonPath
}

$outputs = foreach ($measurement in $measurements) {
    [pscustomobject]@{
        Workload = $measurement.Workload
        Area = $measurement.Definition.Area
        CatalogComparisonId = $executionPlanByWorkload[$measurement.Workload].CatalogComparisonId
        Platform = $platform
        RunMode = $RunMode
        Publish = $executionPlanByWorkload[$measurement.Workload].Publish
        AffinityMask = $affinityLabel
        AffinityApplication = $effectiveAffinityApplication
        QuestPdfPackageVersion = $executionPlanByWorkload[$measurement.Workload].QuestPdfPackageVersion
        InternalOnly = $executionPlanByWorkload[$measurement.Workload].InternalOnly
        SourceCommit = $gitSha
        ArtifactsPath = $measurement.ArtifactsPath
        NormalizedResult = if ($measurement.CatalogEligible) {
            $entry = @($mergedCatalog.Entries | Where-Object {
                $_.ComparisonId -eq $executionPlanByWorkload[$measurement.Workload].CatalogComparisonId -and
                $_.Platform -eq $Platform -and $_.RunMode -eq $RunMode
            })[0]
            Join-Path $staticRoot $entry.ArtifactRelativePath
        } else {
            $measurement.NormalizedResult
        }
        SizeEvidence = $measurement.SizeEvidence
        BundleRoot = $staticRoot
        ComparisonPath = if (-not $InternalQuestPdf) { $comparisonPath } else { $null }
        EvidenceCatalog = if (-not $InternalQuestPdf) { $catalogPath } else { $null }
    }
}

$outputs
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
