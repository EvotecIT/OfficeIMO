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
. (Join-Path $PSScriptRoot 'PSWriteOfficeApiDocs.Bundle.ps1')
. (Join-Path $PSScriptRoot 'PSWriteOfficeApiDocs.Cache.ps1')

function Resolve-RepoRoot {
    param([string] $SiteRootPath, [string] $RequestedRoot)
    $candidate = if ($RequestedRoot) { $RequestedRoot } else {
        Join-Path (Split-Path -Parent $SiteRootPath) 'PowerShell/PSWriteOffice'
    }
    $resolved = (Resolve-Path -LiteralPath $candidate -ErrorAction Stop).Path
    if (-not (Test-Path -LiteralPath (Join-Path $resolved 'PSWriteOffice.psd1') -PathType Leaf)) {
        throw 'The OfficeIMO checkout must contain PowerShell/PSWriteOffice.'
    }
    return $resolved
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

    if ($Source.Repo -cne 'EvotecIT/OfficeIMO' -or [string]::IsNullOrWhiteSpace($Source.Ref)) {
        throw 'PSWriteOffice API source links require the configured repository and a pinned source ref.'
    }
    $metadata = Get-Content -LiteralPath $MetadataPath -Raw | ConvertFrom-Json -Depth 40
    $module = Import-PowerShellDataFile -LiteralPath (Join-Path (Split-Path -Parent $MetadataPath) 'PSWriteOffice.psd1')
    $publicRef = $Source.Ref
    $sourcePrefix = "https://github.com/$($Source.Repo)/blob/$([Uri]::EscapeDataString($publicRef))/"
    $changed = $false
    foreach ($command in @($metadata.commands)) {
        if ([string]::IsNullOrWhiteSpace($command.sourcePath)) { continue }
        $sourcePath = (($command.sourcePath -replace '\\', '/') -split '/' |
            ForEach-Object { [Uri]::EscapeDataString($_) }) -join '/'
        $sourceUrl = $sourcePrefix + 'PowerShell/PSWriteOffice/' + $sourcePath
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
$resolvedRepoRoot = Resolve-RepoRoot -SiteRootPath $resolvedSiteRoot -RequestedRoot $PSWriteOfficeRoot
$sourceHead = (& git -C $resolvedRepoRoot rev-parse --verify HEAD).Trim()
if ($LASTEXITCODE -ne 0) { throw 'Cannot identify the OfficeIMO source revision.' }
$powerShellSource = @([pscustomobject]@{ Repo = 'EvotecIT/OfficeIMO'; Ref = $sourceHead })
if (-not $ExportedArtifactsRoot) {
    Assert-NoSourceLinks -Root $resolvedRepoRoot
    foreach ($snapshotRoot in @('data/apidocs/powershell', 'content/docs/pswriteoffice')) {
        $snapshotPath = Join-Path $resolvedSiteRoot $snapshotRoot
        if (Test-Path -LiteralPath $snapshotPath) { Assert-NoSourceLinks -Root $snapshotPath }
    }
    if (Test-CurrentPSWriteOfficeSnapshot -RepositoryRoot $resolvedRepoRoot -SiteRoot $resolvedSiteRoot `
            -SourceCommit $sourceHead -SkipDocumentation:$SkipDocumentation) {
        $examplesSource = Join-Path $resolvedRepoRoot 'Examples'
        $updateExamples = -not $SkipExamples -and (Test-Path -LiteralPath $examplesSource -PathType Container)
        if ($updateExamples) {
            Sync-DirectoryContents -Source $examplesSource `
                -Destination (Join-Path $resolvedSiteRoot 'data/apidocs/powershell/examples')
        }
        $producer = Get-Content -LiteralPath (Join-Path $resolvedSiteRoot 'data/apidocs/powershell/bundle-manifest.json') -Raw | ConvertFrom-Json
        return [pscustomobject]@{
            siteRoot = $resolvedSiteRoot
            repoRoot = $resolvedRepoRoot
            sourceCommit = $producer.sourceCommit
            validatedSourceCommit = $sourceHead
            apiSnapshotSource = 'verified-unchanged-inputs'
            fallbackUsed = $false
            helpUpdated = $false
            commandMetadataUpdated = $false
            documentationUpdated = $false
            examplesUpdated = $updateExamples
        }
    }
}
if (-not $ExportedArtifactsRoot) {
    $outputRoot = Join-Path ([IO.Path]::GetTempPath()) ('officeimo-powershell-api-' + [guid]::NewGuid().ToString('N'))
    $ExportedArtifactsRoot = Join-Path $outputRoot 'artifacts'
    $null = & (Join-Path $resolvedRepoRoot 'Build/Export-WebsiteArtifacts.ps1') `
        -RepositoryRoot $resolvedRepoRoot -ArtifactsRoot $ExportedArtifactsRoot -BuildOutputRoot $outputRoot -ArtifactsOnly
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

if (-not $sourceBundle.Valid) {
    throw "The current OfficeIMO PowerShell API bundle is invalid: $($sourceBundle.Reason)."
}
# A partial API refresh must retain the evidence for the unchanged documentation snapshot.
$documentationEvidenceBytes = if ($SkipDocumentation -and (Test-Path -LiteralPath $catalogEvidencePath -PathType Leaf)) {
    [IO.File]::ReadAllBytes($catalogEvidencePath)
} else { $null }
Sync-DirectoryContents -Source $websiteArtifactsRoot -Destination $targetRoot
if ($null -ne $documentationEvidenceBytes) {
    [IO.File]::WriteAllBytes($catalogEvidencePath, $documentationEvidenceBytes)
}
# The module help uses a lowercase suffix; the website's API input has a
# canonical uppercase suffix that must also exist on case-sensitive filesystems.
Rename-Item -LiteralPath (Join-Path $targetRoot 'PSWriteOffice-help.xml') -NewName 'PSWriteOffice-Help.xml'
$summary.sourceCommit = $sourceHead
$summary.helpSource = $resolvedHelpPath
$summary.helpUpdated = $true
$summary.manifestSource = $resolvedManifestPath
$summary.manifestUpdated = $true
$summary.commandMetadataSource = $resolvedCommandMetadataPath
$summary.commandMetadataUpdated = $true
$summary.fallbackUsed = $false
$summary.apiSnapshotSource = 'current-source'

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
