param(
    [string] $SiteRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path,
    [string] $PSWriteOfficeRoot = $env:PSWRITEOFFICE_ROOT,
    [string] $ExportedArtifactsRoot,
    [switch] $SkipDocumentation,
    [switch] $SkipExamples
)

if ($ExportedArtifactsRoot -and -not $PSWriteOfficeRoot) {
    throw 'An exported bundle requires an explicit PSWriteOffice source checkout.'
}

function Assert-NoSourceLinks {
    param([Parameter(Mandatory)][string] $Root)

    foreach ($item in @(Get-Item -LiteralPath $Root -Force -ErrorAction Stop) +
                     @(Get-ChildItem -LiteralPath $Root -Recurse -Force -ErrorAction Stop)) {
        if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw "The PSWriteOffice source contains a symbolic link or reparse point: '$($item.FullName)'."
        }
    }
}

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'PSWriteOfficeApiDocs.Catalog.ps1')
. (Join-Path $PSScriptRoot 'PSWriteOfficeApiDocs.Documentation.ps1')

function Resolve-RepoRoot {
    param(
        [string] $SiteRootPath,
        [string] $RequestedRoot
    )

    if (-not [string]::IsNullOrWhiteSpace($RequestedRoot)) {
        try {
            $resolvedRequestedRoot = (Resolve-Path -LiteralPath $RequestedRoot -ErrorAction Stop).Path
        } catch {
            throw "The explicit PSWriteOffice repository root '$RequestedRoot' was not found."
        }

        if (-not (Test-Path -LiteralPath (Join-Path $resolvedRequestedRoot 'PSWriteOffice.psd1') -PathType Leaf)) {
            throw "The explicit PSWriteOffice repository root '$resolvedRequestedRoot' does not contain PSWriteOffice.psd1."
        }

        return $resolvedRequestedRoot
    }

    $candidates = @(
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
        [Parameter(Mandatory)][string[]] $ExpectedCommands,
        [Parameter(Mandatory)][version] $ExpectedVersion,
        [Parameter(Mandatory)][string] $ExpectedSourceCommit
    )

    $bundleManifestPath = Join-Path (Split-Path -Parent $ManifestPath) 'bundle-manifest.json'
    $paths = @($ManifestPath, $HelpPath, $MetadataPath, $bundleManifestPath)
    $missing = @($paths | Where-Object { -not (Test-Path -LiteralPath $_ -PathType Leaf) })
    if ($missing.Count -gt 0) {
        return [PSCustomObject]@{ Valid = $false; Reason = "missing: $($missing -join ', ')"; CommandCount = 0 }
    }

    try {
        $provenance = Get-Content -LiteralPath $bundleManifestPath -Raw | ConvertFrom-Json
        if ($provenance.schemaVersion -ne 1 -or $provenance.moduleName -cne 'PSWriteOffice' -or
            [version]$provenance.moduleVersion -ne $ExpectedVersion -or $provenance.sourceCommit -cne $ExpectedSourceCommit) {
            return [PSCustomObject]@{ Valid = $false; Reason = 'bundle provenance does not match the pinned release'; CommandCount = 0 }
        }
        foreach ($component in @(
            @{ Key = 'manifest'; Path = $ManifestPath },
            @{ Key = 'help'; Path = $HelpPath },
            @{ Key = 'metadata'; Path = $MetadataPath }
        )) {
            $evidence = $provenance.components.($component.Key)
            if ($null -eq $evidence -or (Get-Item -LiteralPath $component.Path).Length -ne $evidence.bytes -or
                (Get-FileHash -LiteralPath $component.Path -Algorithm SHA256).Hash -ine $evidence.sha256) {
                return [PSCustomObject]@{ Valid = $false; Reason = "bundle $($component.Key) does not match its release fingerprint"; CommandCount = 0 }
            }
        }
        $bundleManifest = Import-PowerShellDataFile -LiteralPath $ManifestPath
        if ([version]$bundleManifest.ModuleVersion -ne $ExpectedVersion) {
            return [PSCustomObject]@{ Valid = $false; Reason = "bundle version $($bundleManifest.ModuleVersion) does not match source version $ExpectedVersion"; CommandCount = 0 }
        }
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

function Update-ApiBundleMetadataEvidence {
    param([Parameter(Mandatory)][string] $ApiRoot)

    # Source URLs are normalized after import. Retain the producer's release and
    # other component fingerprints, and record only that derived metadata change.
    $path = Join-Path $ApiRoot 'bundle-manifest.json'
    $bundle = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
    $metadataPath = Join-Path $ApiRoot 'command-metadata.json'
    $bundle.components.metadata.bytes = (Get-Item -LiteralPath $metadataPath).Length
    $bundle.components.metadata.sha256 = (Get-FileHash -LiteralPath $metadataPath -Algorithm SHA256).Hash.ToLowerInvariant()
    $bundle | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $path -Encoding utf8
}

function Set-PSWriteOfficeSourceLinks {
    param(
        [Parameter(Mandatory)][string] $MetadataPath,
        [Parameter(Mandatory)] $Source
    )

    if ($Source.Repo -cne 'EvotecIT/PSWriteOffice' -or [string]::IsNullOrWhiteSpace($Source.Ref)) {
        throw 'PSWriteOffice API source links require the configured repository and a pinned source ref.'
    }
    $metadata = Get-Content -LiteralPath $MetadataPath -Raw | ConvertFrom-Json -Depth 40
    $module = Import-PowerShellDataFile -LiteralPath (Join-Path (Split-Path -Parent $MetadataPath) 'PSWriteOffice.psd1')
    $publicRef = 'v' + [string] $module.ModuleVersion
    $sourcePrefix = "https://github.com/$($Source.Repo)/blob/$([Uri]::EscapeDataString($publicRef))/"
    $changed = $false
    foreach ($command in @($metadata.commands)) {
        if ([string]::IsNullOrWhiteSpace($command.sourcePath)) { continue }
        $sourcePath = (($command.sourcePath -replace '\\', '/') -split '/' |
            ForEach-Object { [Uri]::EscapeDataString($_) }) -join '/'
        $sourceUrl = $sourcePrefix + $sourcePath
        if ([int] $command.sourceLine -gt 0) { $sourceUrl += "#L$($command.sourceLine)" }
        if ($command.sourceUrl -cne $sourceUrl) {
            $command.sourceUrl = $sourceUrl
            $changed = $true
        }
    }
    if ($changed) {
        [IO.File]::WriteAllText($MetadataPath, ($metadata | ConvertTo-Json -Depth 40) + [Environment]::NewLine)
    }
    return $changed
}

$resolvedSiteRoot = (Resolve-Path -LiteralPath $SiteRoot).Path
$siteConfiguration = Get-Content -LiteralPath (Join-Path $resolvedSiteRoot 'site.json') -Raw | ConvertFrom-Json
$powerShellSource = @($siteConfiguration.Sources | Where-Object Slug -CEQ 'pswriteoffice')
if ($powerShellSource.Count -ne 1) { throw 'The website must configure exactly one PSWriteOffice source.' }
$resolvedRepoRoot = Resolve-RepoRoot -SiteRootPath $resolvedSiteRoot -RequestedRoot $PSWriteOfficeRoot
if ($resolvedRepoRoot) {
    $sourceHead = & git -C $resolvedRepoRoot rev-parse --verify HEAD 2>$null
    $headExitCode = $LASTEXITCODE
    $pinnedCommit = & git -C $resolvedRepoRoot rev-parse --verify "$($powerShellSource[0].Ref)^{commit}" 2>$null
    $pinExitCode = $LASTEXITCODE
    $syncManifestPath = Join-Path $resolvedSiteRoot '.powerforge/git-sync-manifest.json'
    if ($pinExitCode -ne 0 -and (Test-Path -LiteralPath $syncManifestPath -PathType Leaf)) {
        # PowerForge can fetch a release into FETCH_HEAD without creating a local tag.
        $syncManifest = Get-Content -LiteralPath $syncManifestPath -Raw | ConvertFrom-Json
        $syncedSource = @($syncManifest.entries | Where-Object {
            $_.repo -ceq "https://github.com/$($powerShellSource[0].Repo).git" -and
            $_.requestedRef -ceq $powerShellSource[0].Ref -and $_.destination -ceq $resolvedRepoRoot
        })
        if ($syncedSource.Count -eq 1 -and $syncedSource[0].resolvedRef -cmatch '^[0-9a-f]{7,64}$') {
            $pinnedCommit = & git -C $resolvedRepoRoot rev-parse --verify "$($syncedSource[0].resolvedRef)^{commit}" 2>$null
            $pinExitCode = $LASTEXITCODE
        }
    }
    $sourceChanges = @(& git -C $resolvedRepoRoot status --porcelain --untracked-files=all 2>$null)
    $statusExitCode = $LASTEXITCODE
    if ($headExitCode -ne 0 -or $pinExitCode -ne 0 -or $statusExitCode -ne 0 -or
        $sourceHead -cne $pinnedCommit -or $sourceChanges.Count -gt 0) {
        if (-not [string]::IsNullOrWhiteSpace($PSWriteOfficeRoot)) {
            throw "The explicit PSWriteOffice source must be clean and checked out at '$($powerShellSource[0].Ref)'."
        }
        Write-Host "The local PSWriteOffice source is not a clean '$($powerShellSource[0].Ref)' checkout. Keeping the checked-in release snapshot." -ForegroundColor Yellow
        $resolvedRepoRoot = $null
    }
}
$targetRoot = Join-Path $resolvedSiteRoot 'data\apidocs\powershell'
$targetHelpPath = Join-Path $targetRoot 'PSWriteOffice-Help.xml'
$targetManifestPath = Join-Path $targetRoot 'PSWriteOffice.psd1'
$targetCommandMetadataPath = Join-Path $targetRoot 'command-metadata.json'
$targetBundleManifestPath = Join-Path $targetRoot 'bundle-manifest.json'
$targetExamplesPath = Join-Path $targetRoot 'examples'
$targetCatalogPath = Join-Path $resolvedSiteRoot 'data/pswriteoffice_command_catalog.json'
$catalogEvidencePath = Join-Path $targetRoot 'documentation-manifest.json'

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
    documentationSource = $null
    documentationUpdated = $false
    documentationNavigationUpdated = $false
    documentationCatalogSource = $null
    documentationCatalogUpdated = $false
    fallbackUsed = $true
    apiSnapshotSource = 'checked-in'
    expectedCommandCount = $null
    sourceBundleStatus = $null
}

if (-not $resolvedRepoRoot) {
    $snapshotCatalog = Get-Content -LiteralPath (Join-Path $resolvedSiteRoot 'data/pswriteoffice_command_catalog.json') -Raw | ConvertFrom-Json
    $snapshotCommands = @(Get-ManifestCommandNames -Path $targetManifestPath)
    $checkedInBundle = Test-ApiBundle -ManifestPath $targetManifestPath -HelpPath $targetHelpPath `
        -MetadataPath $targetCommandMetadataPath -ExpectedCommands $snapshotCommands `
        -ExpectedVersion ([version]$snapshotCatalog.module.version) -ExpectedSourceCommit $powerShellSource[0].Ref
    if (-not $checkedInBundle.Valid -or $snapshotCommands.Count -ne $snapshotCatalog.module.commandCount) {
        throw "The checked-in PowerShell API snapshot is invalid: $($checkedInBundle.Reason)."
    }
    $summary.expectedCommandCount = $snapshotCommands.Count
    if (-not $SkipDocumentation) {
        $catalogEvidence = Assert-DocumentationCatalogSnapshot -CatalogPath $targetCatalogPath -EvidencePath $catalogEvidencePath -SiteRootPath $resolvedSiteRoot `
            -ExpectedSourceRef $powerShellSource[0].Ref -ExpectedVersion ([version]$snapshotCatalog.module.version) -ExpectedCommandCount $snapshotCommands.Count
    }
    $summary.commandMetadataUpdated = Set-PSWriteOfficeSourceLinks -MetadataPath $targetCommandMetadataPath -Source $powerShellSource[0]
    if ($summary.commandMetadataUpdated) { Update-ApiBundleMetadataEvidence -ApiRoot $targetRoot }
    if (-not $SkipDocumentation) {
        $linkUpdates = Set-PSWriteOfficeDocumentationSnapshot -SiteRootPath $resolvedSiteRoot -Source $powerShellSource[0]
        $summary.documentationCatalogUpdated = $linkUpdates.CatalogUpdated
        $summary.documentationUpdated = $linkUpdates.DocumentationUpdated
        Save-DocumentationCatalogEvidence -CatalogPath $targetCatalogPath -EvidencePath $catalogEvidencePath -SiteRootPath $resolvedSiteRoot -Evidence $catalogEvidence
    }
    Write-Host 'No matching PSWriteOffice release checkout found. Keeping checked-in PowerShell API snapshot.' -ForegroundColor Yellow
    [PSCustomObject] $summary
    return
}

# Source documents and examples are copied into publishable site content.
# Reject links before reading any source bundle or dereferencing copy inputs.
Assert-NoSourceLinks -Root $resolvedRepoRoot

$sourceArtifactsRoot = if ($ExportedArtifactsRoot) {
    (Resolve-Path -LiteralPath $ExportedArtifactsRoot -ErrorAction Stop).Path
} else { Join-Path $resolvedRepoRoot 'WebsiteArtifacts' }
if (Test-Path -LiteralPath $sourceArtifactsRoot) {
    Assert-NoSourceLinks -Root $sourceArtifactsRoot
}
$websiteArtifactsRoot = Join-Path $sourceArtifactsRoot 'apidocs\powershell'
$resolvedHelpPath = Join-Path $websiteArtifactsRoot 'PSWriteOffice-help.xml'
$resolvedManifestPath = Join-Path $websiteArtifactsRoot 'PSWriteOffice.psd1'
$resolvedCommandMetadataPath = Join-Path $websiteArtifactsRoot 'command-metadata.json'
$resolvedBundleManifestPath = Join-Path $websiteArtifactsRoot 'bundle-manifest.json'
$authoritativeManifestPath = Join-Path $resolvedRepoRoot 'PSWriteOffice.psd1'
$expectedCommands = @(Get-ManifestCommandNames -Path $authoritativeManifestPath)
$expectedVersion = [version](Import-PowerShellDataFile -LiteralPath $authoritativeManifestPath).ModuleVersion
$summary.expectedCommandCount = $expectedCommands.Count

# Qualify the documentation pair before any API, examples or page writes.
if (-not $SkipDocumentation) {
    $sourceDocumentationPath = Join-Path $resolvedRepoRoot 'Website/content/project-docs/docs'
    $targetDocumentationPath = Join-Path $resolvedSiteRoot 'content/docs/pswriteoffice'
    $sourceDocumentationAvailable = Test-CompleteSourceDocumentation -RepositoryRoot $resolvedRepoRoot -SourceCommit $sourceHead
    if ($sourceDocumentationAvailable) {
        $generatedCatalog = New-DocumentationCatalogSnapshot -RepositoryRoot $resolvedRepoRoot -SourceCommit $sourceHead `
            -SourceRef $powerShellSource[0].Ref -ExpectedVersion $expectedVersion -ExpectedCommandCount $expectedCommands.Count
        $catalogEvidence = $generatedCatalog.Evidence
    } else {
        if (-not (Test-Path -LiteralPath (Join-Path $targetDocumentationPath 'index.md') -PathType Leaf)) {
            throw 'No complete PSWriteOffice source or checked-in documentation snapshot is available.'
        }
        $catalogEvidence = Assert-DocumentationCatalogSnapshot -CatalogPath $targetCatalogPath -EvidencePath $catalogEvidencePath -SiteRootPath $resolvedSiteRoot `
            -ExpectedSourceRef $powerShellSource[0].Ref -ExpectedVersion $expectedVersion -ExpectedCommandCount $expectedCommands.Count
    }
}

$sourceBundle = Test-ApiBundle `
    -ManifestPath $resolvedManifestPath `
    -HelpPath $resolvedHelpPath `
    -MetadataPath $resolvedCommandMetadataPath `
    -ExpectedCommands $expectedCommands -ExpectedVersion $expectedVersion -ExpectedSourceCommit $sourceHead
$summary.sourceBundleStatus = $sourceBundle.Reason

$checkedInBundle = Test-ApiBundle `
    -ManifestPath $targetManifestPath `
    -HelpPath $targetHelpPath `
    -MetadataPath $targetCommandMetadataPath `
    -ExpectedCommands $expectedCommands -ExpectedVersion $expectedVersion -ExpectedSourceCommit $sourceHead

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
    Copy-Item -LiteralPath $resolvedBundleManifestPath -Destination $targetBundleManifestPath -Force
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

$sourceLinksUpdated = Set-PSWriteOfficeSourceLinks -MetadataPath $targetCommandMetadataPath -Source $powerShellSource[0]
if ($sourceLinksUpdated) { Update-ApiBundleMetadataEvidence -ApiRoot $targetRoot }
$summary.commandMetadataUpdated = $summary.commandMetadataUpdated -or $sourceLinksUpdated

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

if (-not $SkipDocumentation) {
    if ($sourceDocumentationAvailable) {
        Sync-DocumentationPages -Source $sourceDocumentationPath -Destination $targetDocumentationPath
        Sync-DocumentationToc -Source $sourceDocumentationPath -TocPath (Join-Path $resolvedSiteRoot 'content/docs/toc.json')
        $summary.documentationNavigationUpdated = $true
        [IO.File]::WriteAllText($targetCatalogPath, $generatedCatalog.Content, [Text.UTF8Encoding]::new($false))
        $summary.documentationSource = $sourceDocumentationPath
        $summary.documentationUpdated = $true
        $summary.documentationCatalogSource = Join-Path $resolvedRepoRoot 'Build/Export-WebsiteDocumentationCatalog.ps1'
        $summary.documentationCatalogUpdated = $true
    } else {
        Write-Host 'PSWriteOffice source documentation is unavailable or incomplete. Keeping the verified checked-in documentation snapshot.' -ForegroundColor Yellow
        $summary.documentationSource = $targetDocumentationPath
        $summary.documentationCatalogSource = $targetCatalogPath
    }
    $linkUpdates = Set-PSWriteOfficeDocumentationSnapshot -SiteRootPath $resolvedSiteRoot -Source $powerShellSource[0]
    $summary.documentationCatalogUpdated = $summary.documentationCatalogUpdated -or $linkUpdates.CatalogUpdated
    $summary.documentationUpdated = $summary.documentationUpdated -or $linkUpdates.DocumentationUpdated
    Save-DocumentationCatalogEvidence -CatalogPath $targetCatalogPath -EvidencePath $catalogEvidencePath -SiteRootPath $resolvedSiteRoot -Evidence $catalogEvidence
}

[PSCustomObject] $summary
