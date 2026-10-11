<#
.SYNOPSIS
Exports the PowerShell API help, command metadata and examples for the website.
.DESCRIPTION
Use an explicit complete module bundle with -SkipBuild to export a qualified package
without building or installing it. The bundle version and help command set must match
the source checkout. -ArtifactsOnly leaves repository documentation unchanged.
.EXAMPLE
./Build/Export-WebsiteArtifacts.ps1 -RepositoryRoot ./release-source -ArtifactsRoot ./out/website-api -SkipBuild -BuiltModuleManifestPath ./package/PSWriteOffice.psd1 -ArtifactsOnly
#>
param(
    [string] $RepositoryRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path,
    [string] $ArtifactsRoot = '',
    [switch] $SkipBuild,
    [string] $BuiltModuleManifestPath = '',
    [switch] $ArtifactsOnly
)

$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($ArtifactsRoot)) {
    $ArtifactsRoot = Join-Path $RepositoryRoot 'WebsiteArtifacts'
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

function Get-RelativeFilePath {
    param(
        [Parameter(Mandatory)][string] $BasePath,
        [Parameter(Mandatory)][string] $Path
    )

    $baseFullPath = [System.IO.Path]::GetFullPath($BasePath)
    $separator = [System.IO.Path]::DirectorySeparatorChar.ToString()
    if (-not $baseFullPath.EndsWith($separator, [System.StringComparison]::Ordinal)) {
        $baseFullPath += $separator
    }

    $targetFullPath = [System.IO.Path]::GetFullPath($Path)
    $baseUri = [System.Uri] $baseFullPath
    $targetUri = [System.Uri] $targetFullPath
    if (-not [string]::Equals($baseUri.Scheme, $targetUri.Scheme, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $targetFullPath
    }

    $relativeUri = $baseUri.MakeRelativeUri($targetUri).ToString()
    return [System.Uri]::UnescapeDataString($relativeUri).Replace('/', [System.IO.Path]::DirectorySeparatorChar)
}

function Get-CompleteModuleBundle {
    param([Parameter(Mandatory)][string] $Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $null }
    $moduleRoot = Split-Path -Parent ([IO.Path]::GetFullPath($Path))
    $helpPath = Join-Path $moduleRoot "en-US\$moduleName-help.xml"
    if (-not (Test-Path -LiteralPath $helpPath -PathType Leaf)) { return $null }
    $builtManifest = Import-PowerShellDataFile -LiteralPath $Path
    if ([version]$builtManifest.ModuleVersion -ne [version]$manifest.ModuleVersion) { return $null }
    $helpCommands = @(Get-HelpCommandNames -Path $helpPath)
    if ($helpCommands.Count -ne $expectedCommands.Count -or
        @(Compare-Object -ReferenceObject $expectedCommands -DifferenceObject $helpCommands).Count -gt 0) {
        return $null
    }
    return [PSCustomObject]@{ HelpPath = $helpPath; ManifestPath = $Path; ModuleRoot = $moduleRoot }
}

function Find-CompleteBuiltModule {
    $unpackedRoot = Join-Path $RepositoryRoot 'Artefacts\Unpacked'
    if (-not (Test-Path -LiteralPath $unpackedRoot -PathType Container)) {
        return $null
    }

    foreach ($helpFile in Get-ChildItem -LiteralPath $unpackedRoot -Recurse -Filter "$moduleName-help.xml" -File |
        Sort-Object LastWriteTimeUtc -Descending) {
        $moduleRoot = Split-Path -Parent (Split-Path -Parent $helpFile.FullName)
        $builtManifestPath = Join-Path $moduleRoot "$moduleName.psd1"
        $bundle = Get-CompleteModuleBundle -Path $builtManifestPath
        if ($bundle) { return $bundle }
    }

    return $null
}

$builtModule = $null
if ($BuiltModuleManifestPath) {
    if (-not $SkipBuild) { throw 'An explicit module bundle requires -SkipBuild.' }
    $builtModule = Get-CompleteModuleBundle -Path $BuiltModuleManifestPath
    if (-not $builtModule) {
        throw "The explicit module must be a complete bundle for PSWriteOffice $($manifest.ModuleVersion), with help for every exported command."
    }
} elseif (-not $SkipBuild) {
    & (Join-Path $RepositoryRoot 'Build\Build-Module.ps1') -RunMode Build
    if (-not $?) {
        throw 'PSWriteOffice module build failed.'
    }
    $builtModule = Find-CompleteBuiltModule
} else {
    $builtModule = Find-CompleteBuiltModule
}

if (-not $builtModule) {
    throw "No complete built PSWriteOffice API bundle was found. Run Build\Build-Module.ps1 -RunMode Build first."
}

$sourceFilesByType = @{}
$sourceFilePriorityByType = @{}
foreach ($sourceFile in Get-ChildItem -LiteralPath (Join-Path $RepositoryRoot 'Sources\PSWriteOffice\Cmdlets') -Recurse -Filter '*.cs' -File) {
    $sourceText = Get-Content -LiteralPath $sourceFile.FullName -Raw -Encoding UTF8
    $containsCmdletDeclaration = $sourceText -match '(?m)^\s*\[(?:System\.Management\.Automation\.)?Cmdlet(?:Attribute)?\s*\('
    $typeNames = @($sourceFile.BaseName) + @(
        [regex]::Matches($sourceText, '\b(?:sealed\s+)?(?:partial\s+)?class\s+(?<name>[A-Za-z_][A-Za-z0-9_]*)\b') |
            ForEach-Object { $_.Groups['name'].Value }
    )

    foreach ($typeName in @($typeNames | Sort-Object -Unique)) {
        $priority = $(if ($containsCmdletDeclaration) { 2 } else { 0 }) + $(if ($sourceFile.BaseName -eq $typeName) { 1 } else { 0 })
        if (-not $sourceFilesByType.ContainsKey($typeName) -or $priority -gt $sourceFilePriorityByType[$typeName]) {
            $sourceFilesByType[$typeName] = $sourceFile.FullName
            $sourceFilePriorityByType[$typeName] = $priority
        }
    }
}

Remove-Module -Name $moduleName -Force -ErrorAction SilentlyContinue
try {
    Import-Module $builtModule.ManifestPath -Force -ErrorAction Stop | Out-Null
    $allCommands = @(Get-Command -Module $moduleName -ErrorAction Stop)
    $commandsByName = @{}
    $aliasesByTarget = @{}

    foreach ($command in $allCommands) {
        $commandsByName[$command.Name] = $command
        if ($command.CommandType -ne 'Alias' -or [string]::IsNullOrWhiteSpace($command.ResolvedCommandName)) {
            continue
        }

        if (-not $aliasesByTarget.ContainsKey($command.ResolvedCommandName)) {
            $aliasesByTarget[$command.ResolvedCommandName] = [System.Collections.Generic.List[string]]::new()
        }
        [void] $aliasesByTarget[$command.ResolvedCommandName].Add($command.Name)
    }

    $missingCommands = @($expectedCommands | Where-Object { -not $commandsByName.ContainsKey($_) })
    if ($missingCommands.Count -gt 0) {
        throw "The built module does not export every manifest cmdlet: $($missingCommands -join ', ')"
    }

    $commandMetadata = foreach ($commandName in $expectedCommands) {
        $command = $commandsByName[$commandName]
        $declaredAliases = @()
        if ($command.ImplementingType) {
            $declaredAliases = @($command.ImplementingType.GetCustomAttributes([System.Management.Automation.AliasAttribute], $true) |
                ForEach-Object { $_.AliasNames })
        }
        $sourcePath = $null
        $sourceLine = $null
        $implementingType = $command.ImplementingType
        if ($implementingType -and $sourceFilesByType.ContainsKey($implementingType.Name)) {
            $sourceFullPath = $sourceFilesByType[$implementingType.Name]
            $sourcePath = (Get-RelativeFilePath -BasePath $RepositoryRoot -Path $sourceFullPath).Replace('\', '/')
            $typePattern = '\bclass\s+' + [regex]::Escape($implementingType.Name) + '\b'
            $typeMatch = Select-String -LiteralPath $sourceFullPath -Pattern $typePattern | Select-Object -First 1
            if ($typeMatch) {
                $sourceLine = $typeMatch.LineNumber
            }
        }

        $entry = [ordered]@{
            name = $commandName
            kind = 'Cmdlet'
            aliases = @(@($aliasesByTarget[$commandName]) + $declaredAliases |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                Sort-Object -Unique)
        }
        if ($sourcePath) {
            $entry.sourcePath = $sourcePath
            $entry.sourceLine = $sourceLine
            $entry.sourceUrl = "https://github.com/EvotecIT/PSWriteOffice/blob/main/$sourcePath#L$sourceLine"
        }
        $entry
    }
} finally {
    Remove-Module -Name $moduleName -Force -ErrorAction SilentlyContinue
}

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
$sourceCommitLines = @(& git -C $RepositoryRoot rev-parse HEAD 2>$null)
$sourceCommit = if ($LASTEXITCODE -eq 0 -and $sourceCommitLines.Count -eq 1) { [string]$sourceCommitLines[0] } else { $null }
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
    moduleVersion = [string]$manifest.ModuleVersion; sourceCommit = $sourceCommit
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
}
