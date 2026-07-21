param(
    [string] $SiteRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path,
    [string] $PSWriteOfficeRoot = '',
    [switch] $SkipExamples
)

$ErrorActionPreference = 'Stop'

function Resolve-RepoRoot {
    param(
        [string] $SiteRootPath,
        [string] $RequestedRoot
    )

    $candidates = @()
    if (-not [string]::IsNullOrWhiteSpace($RequestedRoot)) {
        $candidates += $RequestedRoot
    }

    $candidates += @(
        (Join-Path $SiteRootPath 'projects\pswriteoffice'),
        (Join-Path (Split-Path -Parent $SiteRootPath) 'PSWriteOffice'),
        'C:\Support\GitHub\PSWriteOffice',
        '/mnt/c/Support/GitHub/PSWriteOffice'
    )

    foreach ($candidate in ($candidates | Select-Object -Unique)) {
        if ([string]::IsNullOrWhiteSpace($candidate)) {
            continue
        }

        try {
            $resolved = (Resolve-Path -LiteralPath $candidate -ErrorAction Stop).Path
        } catch {
            continue
        }

        if (Test-Path -LiteralPath (Join-Path $resolved 'PSWriteOffice.psd1') -PathType Leaf) {
            return $resolved
        }
    }

    return $null
}

function Sync-DirectoryContents {
    param(
        [Parameter(Mandatory)]
        [string] $Source,
        [Parameter(Mandatory)]
        [string] $Destination
    )

    New-Item -ItemType Directory -Path $Destination -Force | Out-Null

    $existing = Get-ChildItem -LiteralPath $Destination -Recurse -Force -ErrorAction SilentlyContinue |
        Sort-Object FullName -Descending
    foreach ($item in $existing) {
        if ($item.Name -eq '.gitkeep') {
            continue
        }

        Remove-Item -LiteralPath $item.FullName -Force -Recurse -ErrorAction SilentlyContinue
    }

    $sourceItems = Get-ChildItem -LiteralPath $Source -Recurse -Force -ErrorAction SilentlyContinue
    foreach ($item in $sourceItems) {
        $relativePath = [System.IO.Path]::GetRelativePath($Source, $item.FullName)
        if ([string]::IsNullOrWhiteSpace($relativePath) -or $relativePath -eq '.') {
            continue
        }

        $targetPath = Join-Path $Destination $relativePath
        if ($item.PSIsContainer) {
            New-Item -ItemType Directory -Path $targetPath -Force | Out-Null
            continue
        }

        New-Item -ItemType Directory -Path (Split-Path -Parent $targetPath) -Force | Out-Null
        Copy-Item -LiteralPath $item.FullName -Destination $targetPath -Force
    }
}

function Get-ManifestCommandNames {
    param([Parameter(Mandatory)][string] $Path)

    $manifest = Import-PowerShellDataFile -LiteralPath $Path
    @($manifest.FunctionsToExport) + @($manifest.CmdletsToExport) |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and $_ -ne '*' } |
        Sort-Object -Unique
}

function Get-HelpCommandNames {
    param([Parameter(Mandatory)][string] $Path)

    [xml] $help = Get-Content -LiteralPath $Path -Raw
    @($help.helpItems.command) |
        ForEach-Object { [string] $_.details.name } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Sort-Object -Unique
}

function Get-MetadataCommandNames {
    param([Parameter(Mandatory)][string] $Path)

    $metadata = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
    @($metadata.commands) |
        ForEach-Object { [string] $_.name } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Sort-Object -Unique
}

function Test-ApiBundle {
    param(
        [Parameter(Mandatory)][string] $ManifestPath,
        [Parameter(Mandatory)][string] $HelpPath,
        [Parameter(Mandatory)][string] $MetadataPath,
        [Parameter(Mandatory)][string[]] $ExpectedCommands
    )

    $paths = @($ManifestPath, $HelpPath, $MetadataPath)
    $missing = @($paths | Where-Object { -not (Test-Path -LiteralPath $_ -PathType Leaf) })
    if ($missing.Count -gt 0) {
        return [PSCustomObject]@{ Valid = $false; Reason = "missing: $($missing -join ', ')"; CommandCount = 0 }
    }

    try {
        $sets = [ordered]@{
            manifest = @(Get-ManifestCommandNames -Path $ManifestPath)
            help = @(Get-HelpCommandNames -Path $HelpPath)
            metadata = @(Get-MetadataCommandNames -Path $MetadataPath)
        }
    } catch {
        return [PSCustomObject]@{ Valid = $false; Reason = "could not parse bundle: $($_.Exception.Message)"; CommandCount = 0 }
    }

    foreach ($entry in $sets.GetEnumerator()) {
        $difference = @(Compare-Object -ReferenceObject $ExpectedCommands -DifferenceObject $entry.Value)
        if ($difference.Count -gt 0) {
            return [PSCustomObject]@{
                Valid = $false
                Reason = "$($entry.Key) covers $($entry.Value.Count) of $($ExpectedCommands.Count) authoritative commands"
                CommandCount = $entry.Value.Count
            }
        }
    }

    [PSCustomObject]@{ Valid = $true; Reason = 'complete'; CommandCount = $ExpectedCommands.Count }
}

$resolvedSiteRoot = (Resolve-Path -LiteralPath $SiteRoot).Path
$resolvedRepoRoot = Resolve-RepoRoot -SiteRootPath $resolvedSiteRoot -RequestedRoot $PSWriteOfficeRoot
$targetRoot = Join-Path $resolvedSiteRoot 'data\apidocs\powershell'
$targetHelpPath = Join-Path $targetRoot 'PSWriteOffice-Help.xml'
$targetManifestPath = Join-Path $targetRoot 'PSWriteOffice.psd1'
$targetCommandMetadataPath = Join-Path $targetRoot 'command-metadata.json'
$targetExamplesPath = Join-Path $targetRoot 'examples'

New-Item -ItemType Directory -Path $targetRoot -Force | Out-Null
New-Item -ItemType Directory -Path $targetExamplesPath -Force | Out-Null

$summary = [ordered]@{
    siteRoot = $resolvedSiteRoot
    repoRoot = $resolvedRepoRoot
    sourceCommit = $null
    helpSource = $null
    helpUpdated = $false
    manifestSource = $null
    manifestUpdated = $false
    commandMetadataSource = $null
    commandMetadataUpdated = $false
    examplesSource = $null
    examplesUpdated = $false
    fallbackUsed = $true
    apiSnapshotSource = 'checked-in'
    expectedCommandCount = $null
    sourceBundleStatus = $null
}

if (-not $resolvedRepoRoot) {
    Write-Host 'PSWriteOffice repo not found. Keeping checked-in PowerShell API snapshot.' -ForegroundColor Yellow
    [PSCustomObject] $summary
    return
}

$websiteArtifactsRoot = Join-Path $resolvedRepoRoot 'WebsiteArtifacts\apidocs\powershell'
$resolvedHelpPath = Join-Path $websiteArtifactsRoot 'PSWriteOffice-help.xml'
$resolvedManifestPath = Join-Path $websiteArtifactsRoot 'PSWriteOffice.psd1'
$resolvedCommandMetadataPath = Join-Path $websiteArtifactsRoot 'command-metadata.json'
$authoritativeManifestPath = Join-Path $resolvedRepoRoot 'PSWriteOffice.psd1'
$expectedCommands = @(Get-ManifestCommandNames -Path $authoritativeManifestPath)
$summary.expectedCommandCount = $expectedCommands.Count

$sourceBundle = Test-ApiBundle `
    -ManifestPath $resolvedManifestPath `
    -HelpPath $resolvedHelpPath `
    -MetadataPath $resolvedCommandMetadataPath `
    -ExpectedCommands $expectedCommands
$summary.sourceBundleStatus = $sourceBundle.Reason

$checkedInBundle = Test-ApiBundle `
    -ManifestPath $targetManifestPath `
    -HelpPath $targetHelpPath `
    -MetadataPath $targetCommandMetadataPath `
    -ExpectedCommands $expectedCommands

$gitDirectory = Join-Path $resolvedRepoRoot '.git'
if (Test-Path -LiteralPath $gitDirectory) {
    $summary.sourceCommit = (& git -C $resolvedRepoRoot rev-parse HEAD 2>$null | Select-Object -First 1)
}

if ($sourceBundle.Valid) {
    Copy-Item -LiteralPath $resolvedHelpPath -Destination $targetHelpPath -Force
    $summary.helpSource = $resolvedHelpPath
    $summary.helpUpdated = $true

    Copy-Item -LiteralPath $resolvedManifestPath -Destination $targetManifestPath -Force
    $summary.manifestSource = $resolvedManifestPath
    $summary.manifestUpdated = $true

    Copy-Item -LiteralPath $resolvedCommandMetadataPath -Destination $targetCommandMetadataPath -Force
    $summary.commandMetadataSource = $resolvedCommandMetadataPath
    $summary.commandMetadataUpdated = $true
    $summary.fallbackUsed = $false
    $summary.apiSnapshotSource = 'source-bundle'
} elseif ($checkedInBundle.Valid) {
    Write-Host "PSWriteOffice WebsiteArtifacts are stale ($($sourceBundle.Reason)). Keeping the complete checked-in $($expectedCommands.Count)-command snapshot." -ForegroundColor Yellow
    $summary.helpSource = $targetHelpPath
    $summary.manifestSource = $targetManifestPath
    $summary.commandMetadataSource = $targetCommandMetadataPath
} else {
    throw "No complete PSWriteOffice API snapshot is available. Source bundle: $($sourceBundle.Reason). Checked-in bundle: $($checkedInBundle.Reason)."
}

if (-not $SkipExamples) {
    $sourceExamplesPath = Join-Path $resolvedRepoRoot 'Examples'
    if (Test-Path -LiteralPath $sourceExamplesPath -PathType Container) {
        Sync-DirectoryContents -Source $sourceExamplesPath -Destination $targetExamplesPath
        $summary.examplesSource = $sourceExamplesPath
        $summary.examplesUpdated = $true
    } else {
        Write-Host 'PSWriteOffice examples folder not found in synced repo. Keeping checked-in fallback examples.' -ForegroundColor Yellow
    }
}

[PSCustomObject] $summary
