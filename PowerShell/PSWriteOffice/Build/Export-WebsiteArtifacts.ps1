<#
.SYNOPSIS
Exports the PowerShell API help, command metadata and examples for the website.
.DESCRIPTION
Builds an unsigned module from committed source using PowerForge, without installing it.
Exports only the bundle returned by that build, from a new staging directory. Both host
assemblies, the manifest and help must match the source. -ArtifactsOnly limits the export
writes to ArtifactsRoot; the native build still regenerates Docs. BuildOutputRoot holds
the retained build staging directory. Command metadata uses a clean child process of the
same PowerShell host. BuildOutputRoot and custom ArtifactsRoot must be outside the
source checkout. The default WebsiteArtifacts target refreshes the tracked snapshot;
commit that output before another export. Requires PSPublishModule 3.0.168 or newer.
.EXAMPLE
./Build/Export-WebsiteArtifacts.ps1 -RepositoryRoot ./release-source -ArtifactsRoot ../website-api -BuildOutputRoot ../website-builds -ArtifactsOnly
#>
param(
    [string] $RepositoryRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path,
    [string] $ArtifactsRoot = '',
    [string] $BuildOutputRoot = [System.IO.Path]::GetTempPath(),
    [switch] $ArtifactsOnly
)

$ErrorActionPreference = 'Stop'
$RepositoryRoot = (Resolve-Path -LiteralPath $RepositoryRoot -ErrorAction Stop).ProviderPath

if ([string]::IsNullOrWhiteSpace($ArtifactsRoot)) {
    $ArtifactsRoot = Join-Path $RepositoryRoot 'WebsiteArtifacts'
}
$ArtifactsRoot = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ArtifactsRoot)
$buildRoot = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($BuildOutputRoot)
$pathComparison = if ([Environment]::OSVersion.Platform -eq [PlatformID]::Win32NT) {
    [StringComparison]::OrdinalIgnoreCase
} else { [StringComparison]::Ordinal }
$sourceBoundary = $RepositoryRoot.TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar
foreach ($output in @(@{Name='BuildOutputRoot';Path=$buildRoot}, @{Name='ArtifactsRoot';Path=$ArtifactsRoot})) {
    $insideSource = [string]::Equals($output.Path.TrimEnd('\', '/'), $RepositoryRoot.TrimEnd('\', '/'), $pathComparison) -or
        $output.Path.StartsWith($sourceBoundary, $pathComparison)
    $trackedSnapshot = $output.Name -eq 'ArtifactsRoot' -and
        [string]::Equals($output.Path.TrimEnd('\', '/'), (Join-Path $RepositoryRoot 'WebsiteArtifacts'), $pathComparison)
    if ($insideSource -and -not $trackedSnapshot) {
        throw "$($output.Name) must be outside the source checkout; ArtifactsRoot may also use the tracked WebsiteArtifacts target."
    }
}

$moduleName = 'PSWriteOffice'
$manifestPath = Join-Path $RepositoryRoot "$moduleName.psd1"
$sourceManifestBytes = [System.IO.File]::ReadAllBytes($manifestPath)
$manifest = Import-PowerShellDataFile -LiteralPath $manifestPath
$expectedCommands = @($manifest.CmdletsToExport) |
    Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and $_ -ne '*' } |
    Sort-Object -Unique

function Get-HelpCommandNames {
    param([Parameter(Mandatory)][string] $Path)

    [xml] $help = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
    @($help.helpItems.command) |
        ForEach-Object { [string] $_.details.name } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Sort-Object -Unique
}

function ConvertTo-ComparableManifestValue {
    param([AllowNull()] $Value)

    if ($Value -is [System.Collections.IDictionary]) {
        $result = [ordered]@{}
        foreach ($key in @($Value.Keys | Sort-Object)) {
            $result[$key] = ConvertTo-ComparableManifestValue $Value[$key]
        }
        return $result
    }
    if ($Value -is [array]) {
        return ,@($Value | ForEach-Object { ConvertTo-ComparableManifestValue $_ })
    }
    return $Value
}

function Get-ComparableManifest {
    param([System.Collections.IDictionary] $Data, [switch] $IgnoreVersion)

    $copy = @{}
    foreach ($key in $Data.Keys) {
        if (-not $IgnoreVersion -or $key -ne 'ModuleVersion') { $copy[$key] = $Data[$key] }
    }
    ConvertTo-ComparableManifestValue $copy | ConvertTo-Json -Depth 30 -Compress
}

function Get-CompleteModuleBundle {
    param([Parameter(Mandatory)][string] $Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $null }
    $resolvedManifestPath = (Resolve-Path -LiteralPath $Path -ErrorAction Stop).ProviderPath
    $moduleRoot = Split-Path -Parent $resolvedManifestPath
    $helpPath = Join-Path $moduleRoot "en-US\$moduleName-help.xml"
    if (-not (Test-Path -LiteralPath $helpPath -PathType Leaf)) { return $null }
    $builtManifest = Import-PowerShellDataFile -LiteralPath $Path
    if ([version]$builtManifest.ModuleVersion -ne [version]$manifest.ModuleVersion) { return $null }
    if ((Get-ComparableManifest $builtManifest) -cne (Get-ComparableManifest $manifest)) { return $null }
    $helpCommands = @(Get-HelpCommandNames -Path $helpPath)
    if ($helpCommands.Count -ne $expectedCommands.Count -or
        @(Compare-Object -ReferenceObject $expectedCommands -DifferenceObject $helpCommands).Count -gt 0) {
        return $null
    }
    $expectedAliases = @($manifest.AliasesToExport | Sort-Object -Unique)
    $bundleAliases = @($builtManifest.AliasesToExport | Sort-Object -Unique)
    if ($expectedAliases.Count -ne $bundleAliases.Count -or
        ($expectedAliases.Count -gt 0 -and @(Compare-Object $expectedAliases $bundleAliases).Count -gt 0)) {
        return $null
    }
    return [PSCustomObject]@{ HelpPath = $helpPath; ManifestPath = $resolvedManifestPath; ModuleRoot = $moduleRoot }
}

function Invoke-SourceGit {
    param([Parameter(Mandatory)][string[]] $Arguments)
    # Windows PowerShell turns native stderr into errors before callers can
    # classify Git's exit code. Keep that handling local to these read-only probes.
    $ErrorActionPreference = 'Continue'
    $output = @(& git -C $RepositoryRoot @Arguments 2>$null)
    [PSCustomObject]@{ ExitCode = $LASTEXITCODE; Output = $output }
}

function Get-QualifiedSourceRevision {
    param([string] $ModuleRoot, [switch] $BeforeBuild)

    $topLevelResult = Invoke-SourceGit @('rev-parse', '--show-toplevel')
    $topLevel = $topLevelResult.Output
    $pathComparison = if ([Environment]::OSVersion.Platform -eq [PlatformID]::Win32NT) {
        [StringComparison]::OrdinalIgnoreCase
    } else { [StringComparison]::Ordinal }
    $expectedModuleRoot = if ($topLevelResult.ExitCode -eq 0 -and $topLevel.Count -eq 1) {
        Join-Path $topLevel[0] 'PowerShell/PSWriteOffice'
    }
    if ($topLevelResult.ExitCode -ne 0 -or $topLevel.Count -ne 1 -or
        -not [string]::Equals([IO.Path]::GetFullPath($expectedModuleRoot).TrimEnd('\', '/'),
            $RepositoryRoot.TrimEnd('\', '/'), $pathComparison)) {
        throw 'Website API export requires the OfficeIMO PowerShell module source root.'
    }
    $headResult = Invoke-SourceGit @('rev-parse', '--verify', 'HEAD')
    $head = $headResult.Output
    if ($headResult.ExitCode -ne 0 -or $head.Count -ne 1 -or $head[0] -cnotmatch '^[0-9a-f]{40}$') {
        throw 'Website API export requires a committed source revision.'
    }

    # Provenance starts with a fresh build from committed source. SDK metadata is
    # an additional consistency check, not evidence for an arbitrary input bundle.
    $hostDirectories = @('Core'; if ([Environment]::OSVersion.Platform -eq [PlatformID]::Win32NT) { 'Default' })
    $packageCommits = @(if ($BeforeBuild) { $head[0] } else { foreach ($hostDirectory in $hostDirectories) {
        $assemblyPath = Join-Path $ModuleRoot "Lib/$hostDirectory/PSWriteOffice.dll"
        if (-not (Test-Path -LiteralPath $assemblyPath -PathType Leaf)) {
            throw "The module bundle lacks its $hostDirectory assembly source provenance."
        }
        $information = [Diagnostics.FileVersionInfo]::GetVersionInfo($assemblyPath).ProductVersion
        if ($information -cnotmatch ('^' + [regex]::Escape([string]$manifest.ModuleVersion) + '\+([0-9a-f]{40})$')) {
            throw "The $hostDirectory assembly must identify its module version and source revision."
        }
        $Matches[1]
    } })
    if (@($packageCommits | Sort-Object -Unique).Count -ne 1) {
        throw 'The module host assemblies must come from the same source revision.'
    }
    if ($packageCommits[0] -cne $head[0]) {
        throw 'The freshly built module must identify the selected source revision.'
    }
    $packageRevision = Invoke-SourceGit @('cat-file', '-e', "$($packageCommits[0])^{commit}")
    if ($packageRevision.ExitCode -ne 0) {
        throw "Package source revision $($packageCommits[0]) is unavailable locally. Fetch this revision or deepen the checkout before exporting."
    }
    foreach ($manifestRevision in @($packageCommits[0], $head[0], '') | Select-Object -Unique) {
        $packageManifestResult = Invoke-SourceGit @('show', "${manifestRevision}:${sourcePrefix}PSWriteOffice.psd1")
        if ($packageManifestResult.ExitCode -ne 0) { throw 'A recorded source revision lacks its module manifest.' }
        $parseErrors = $null
        $tokens = $null
        $manifestAst = [System.Management.Automation.Language.Parser]::ParseInput(
            ($packageManifestResult.Output -join "`n").TrimStart([char]0xfeff), [ref]$tokens, [ref]$parseErrors)
        if ($parseErrors.Count -gt 0) { throw 'The package source manifest could not be parsed safely.' }
        $manifestExpression = if ($manifestAst.EndBlock.Statements.Count -eq 1 -and
            $manifestAst.EndBlock.Statements[0] -is [System.Management.Automation.Language.PipelineAst]) {
            $manifestAst.EndBlock.Statements[0].GetPureExpression()
        }
        if ($manifestExpression -isnot [System.Management.Automation.Language.HashtableAst]) {
            throw 'The package source manifest must contain a data hashtable.'
        }
        $packageManifest = $manifestExpression.SafeGetValue()
        # PowerForge resolves ModuleVersion during a normal build; every other field
        # must still describe the manifest from which the package was built.
        if ((Get-ComparableManifest $packageManifest -IgnoreVersion) -cne
            (Get-ComparableManifest $manifest -IgnoreVersion)) {
            throw 'The module bundle source manifest does not match this checkout. Rebuild it from this checkout.'
        }
    }
    $apiInputs = @('Sources', 'PSWriteOffice.psm1', 'Directory.Build.props', 'Directory.Build.targets',
        'Directory.Packages.props', 'global.json', 'NuGet.Config', 'nuget.config', 'Build')
    if ($BeforeBuild) { $apiInputs = @('.', ':(exclude)PSWriteOffice.psd1') }
    $difference = Invoke-SourceGit (@('diff', '--quiet', $packageCommits[0], $head[0], '--') + $apiInputs)
    $workingDifference = Invoke-SourceGit (@('diff', '--quiet', 'HEAD', '--') + $apiInputs)
    $indexDifference = Invoke-SourceGit (@('diff', '--cached', '--quiet', 'HEAD', '--') + $apiInputs)
    $untrackedInputs = Invoke-SourceGit (@('ls-files', '--others', '--exclude-standard', '--') + $apiInputs)
    if ($difference.ExitCode -ne 0 -or $workingDifference.ExitCode -ne 0 -or $indexDifference.ExitCode -ne 0 -or
        $untrackedInputs.ExitCode -ne 0 -or $untrackedInputs.Output.Count -gt 0) {
        throw 'The module bundle source does not match the selected API and build inputs. Rebuild it from this checkout.'
    }
    # The module now compiles the engines from the same checkout. Their source and
    # shared build settings are part of its provenance, even outside the module folder.
    foreach ($arguments in @(
        @('diff', '--quiet', 'HEAD', '--'),
        @('diff', '--cached', '--quiet', 'HEAD', '--')
    )) {
        & git -C $topLevel[0] @arguments 'OfficeIMO.*' 'Directory.Build.*' 'global.json' 'Build'
        if ($LASTEXITCODE -ne 0) {
            throw 'The module bundle source does not match the selected API and build inputs. Commit the engine and shared build inputs first.'
        }
    }
    [PSCustomObject]@{ SourceCommit = [string]$head[0]; PackageSourceCommit = [string]$packageCommits[0] }
}

$sourcePrefix = 'PowerShell/PSWriteOffice/'
$sourceBeforeBuild = Get-QualifiedSourceRevision -BeforeBuild
$stagingPath = Join-Path $buildRoot ('PSWriteOffice-website-export-' + [guid]::NewGuid().ToString('N'))
$buildResults = @(& (Join-Path $RepositoryRoot 'Build/Build-Module.ps1') `
    -RunMode Build -SignModule $false -IncludePackages $false -SkipInstall -PassThru -StagingPath $stagingPath)
if (-not $? -or $buildResults.Count -ne 1 -or -not $buildResults[0].BuildResult) {
    throw 'PSWriteOffice module build did not return one completed build result.'
}
$buildResult = $buildResults[0].BuildResult
$resolvedStaging = (Resolve-Path -LiteralPath $buildResult.StagingPath).ProviderPath
if (-not [string]::Equals($resolvedStaging.TrimEnd('\', '/'), $stagingPath.TrimEnd('\', '/'), [StringComparison]::OrdinalIgnoreCase)) {
    throw 'PSWriteOffice module build returned a different staging directory.'
}
$sourceManifestBytes = [System.IO.File]::ReadAllBytes($manifestPath)
$manifest = Import-PowerShellDataFile -LiteralPath $manifestPath
$expectedCommands = @($manifest.CmdletsToExport) |
    Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and $_ -ne '*' } | Sort-Object -Unique
$builtModule = Get-CompleteModuleBundle -Path $buildResult.ManifestPath
if (-not $builtModule) { throw 'The fresh module build does not contain matching manifest and complete command help.' }
$sourceRevision = Get-QualifiedSourceRevision -ModuleRoot $builtModule.ModuleRoot
if ($sourceBeforeBuild.SourceCommit -cne $sourceRevision.SourceCommit) {
    throw 'The source revision changed during the website API build.'
}

# Windows PowerShell cannot unload dependencies imported by the build tool.
# Use PowerForge's runner to inspect the fresh module in the same host, in a clean process.
$metadataPath = Join-Path $resolvedStaging 'website-command-metadata.json'
$hostExecutable = (Get-Process -Id $PID).Path
$metadataArguments = [string[]]@('-RepositoryRoot', $RepositoryRoot, '-ManifestPath', $builtModule.ManifestPath,
    '-OutputPath', $metadataPath)
# PowerForge is loaded in PSPublishModule's assembly context on Core; resolve its public
# runner types from the returned build result rather than loading another assembly copy.
$powerForgeAssembly = $buildResult.GetType().Assembly
$requestType = $powerForgeAssembly.GetType('PowerForge.PowerShellRunRequest', $true)
$request = $requestType::new((Join-Path $RepositoryRoot 'Build/Export-WebsiteCommandMetadata.ps1'), $metadataArguments,
        [TimeSpan]::FromMinutes(5), ($PSVersionTable.PSEdition -eq 'Core'), $RepositoryRoot,
        $null, $hostExecutable, $true, $true)
$runnerType = $powerForgeAssembly.GetType('PowerForge.PowerShellRunner', $true)
$runner = $runnerType::new($null)
$metadataResult = $runner.Run($request)
if ($metadataResult.ExitCode -ne 0 -or -not (Test-Path -LiteralPath $metadataPath -PathType Leaf)) {
    throw "Command metadata discovery failed in $($metadataResult.Executable): $($metadataResult.StdErr) $($metadataResult.StdOut)"
}
$commandMetadata = Get-Content -LiteralPath $metadataPath -Raw -Encoding UTF8 | ConvertFrom-Json
if (@($commandMetadata).Count -ne $expectedCommands.Count) { throw 'Command metadata count does not match the source manifest.' }

$apiRoot = Join-Path $ArtifactsRoot 'apidocs\powershell'
New-Item -ItemType Directory -Path $apiRoot -Force | Out-Null

$sourceExamplesRoot = Join-Path $RepositoryRoot 'Examples'
$apiExamplesRoot = Join-Path $apiRoot 'examples'
$artifactsBoundary = [System.IO.Path]::GetFullPath($ArtifactsRoot).TrimEnd(
    [System.IO.Path]::DirectorySeparatorChar,
    [System.IO.Path]::AltDirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar
$resolvedExamplesRoot = [System.IO.Path]::GetFullPath($apiExamplesRoot)
if (-not $resolvedExamplesRoot.StartsWith($artifactsBoundary, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "Refusing to refresh examples outside the website artifacts root: $resolvedExamplesRoot"
}
if (Test-Path -LiteralPath $apiExamplesRoot -PathType Container) {
    [System.IO.Directory]::Delete($resolvedExamplesRoot, $true)
}
New-Item -ItemType Directory -Path $apiExamplesRoot -Force | Out-Null
Get-ChildItem -LiteralPath $sourceExamplesRoot -Force |
    Copy-Item -Destination $apiExamplesRoot -Recurse -Force

$apiHelpPath = Join-Path $apiRoot "$moduleName-help.xml"
$docsHelpPath = Join-Path $RepositoryRoot "Docs\Generated\$moduleName-help.xml"
Copy-Item -LiteralPath $builtModule.HelpPath -Destination $apiHelpPath -Force
if (-not $ArtifactsOnly) {
    New-Item -ItemType Directory -Path (Split-Path -Parent $docsHelpPath) -Force | Out-Null
    Copy-Item -LiteralPath $builtModule.HelpPath -Destination $docsHelpPath -Force
}
[System.IO.File]::WriteAllBytes((Join-Path $apiRoot "$moduleName.psd1"), $sourceManifestBytes)

$metadata = [ordered]@{
    schemaVersion = 1
    moduleName = $moduleName
    moduleVersion = [string]$manifest.ModuleVersion
    commands = @($commandMetadata)
}
$metadata | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $apiRoot 'command-metadata.json') -Encoding utf8

# Bind the generated components together so a consumer can detect a mixed release
# even when old and current help happen to contain the same command names.
$components = [ordered]@{}
foreach ($component in @(
    @{ Key = 'manifest'; Path = (Join-Path $apiRoot "$moduleName.psd1") },
    @{ Key = 'help'; Path = $apiHelpPath },
    @{ Key = 'metadata'; Path = (Join-Path $apiRoot 'command-metadata.json') }
)) {
    $components[$component.Key] = [ordered]@{
        bytes = (Get-Item -LiteralPath $component.Path).Length
        sha256 = (Get-FileHash -LiteralPath $component.Path -Algorithm SHA256).Hash.ToLowerInvariant()
    }
}
[ordered]@{
    schemaVersion = 1; moduleName = $moduleName
    moduleVersion = [string]$manifest.ModuleVersion; sourceCommit = $sourceRevision.SourceCommit
    packageSourceCommit = $sourceRevision.PackageSourceCommit
    components = $components
} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $apiRoot 'bundle-manifest.json') -Encoding utf8

& (Join-Path $PSScriptRoot 'Export-WebsiteDocumentationCatalog.ps1') `
    -RepositoryRoot $RepositoryRoot `
    -ManifestPath (Join-Path $apiRoot "$moduleName.psd1") `
    -OutputPath (Join-Path $ArtifactsRoot 'documentation\command-catalog.json') | Out-Null

$catalog = Get-Content -LiteralPath (Join-Path $ArtifactsRoot 'documentation\command-catalog.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$projectManifestPath = Join-Path $ArtifactsRoot 'project-manifest.json'
if (Test-Path -LiteralPath $projectManifestPath -PathType Leaf) {
    $projectManifest = Get-Content -LiteralPath $projectManifestPath -Raw -Encoding UTF8
    $projectManifest = [regex]::Replace($projectManifest, '("version"\s*:\s*")[^"]+("\s*[,}])', ('${1}' + $manifest.ModuleVersion + '${2}'))
    [System.IO.File]::WriteAllText($projectManifestPath, $projectManifest, [System.Text.UTF8Encoding]::new($false))
}

foreach ($guideName in $(if ($ArtifactsOnly) { @() } else { @('command-families.md', 'overview.md') })) {
    $guidePath = Join-Path $RepositoryRoot "Website\content\project-docs\docs\$guideName"
    if (-not (Test-Path -LiteralPath $guidePath -PathType Leaf)) {
        continue
    }
    $guide = Get-Content -LiteralPath $guidePath -Raw -Encoding UTF8
    foreach ($family in $catalog.families) {
        $title = [regex]::Escape([string] $family.title)
        $guide = [regex]::Replace($guide, "(?m)^(\|\s*$title\s*\|\s*)\d+(\s*\|)", ('${1}' + $family.commandCount + '${2}'))
        $guide = [regex]::Replace($guide, "(?m)^(-\s+\*\*$title\s+\u2014\s+)\d+(\s+commands:)", ('${1}' + $family.commandCount + '${2}'))
    }
    [System.IO.File]::WriteAllText($guidePath, $guide, [System.Text.UTF8Encoding]::new($false))
}

[PSCustomObject]@{
    OutputPath = (Resolve-Path -LiteralPath $apiRoot).Path
    CommandCount = $commandMetadata.Count
    SourceMappedCount = @($commandMetadata | Where-Object sourcePath).Count
    HelpPath = $builtModule.HelpPath
    BuildStagingPath = $resolvedStaging
}
