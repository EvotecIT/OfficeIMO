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
        (Join-Path (Split-Path -Parent $SiteRootPath) 'PSWriteOffice'),
        'C:\Support\GitHub\PSWriteOffice',
        '/mnt/c/Support/GitHub/PSWriteOffice',
        (Join-Path $SiteRootPath 'projects\pswriteoffice')
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
    helpSource = $null
    helpUpdated = $false
    manifestSource = $null
    manifestUpdated = $false
    commandMetadataSource = $null
    commandMetadataUpdated = $false
    examplesSource = $null
    examplesUpdated = $false
    fallbackUsed = $true
}

if (-not $resolvedRepoRoot) {
    Write-Host 'PSWriteOffice repo not found. Keeping checked-in PowerShell API snapshot.' -ForegroundColor Yellow
    [PSCustomObject] $summary
    return
}

$helpCandidates = @(
    (Join-Path $resolvedRepoRoot 'WebsiteArtifacts\apidocs\powershell\PSWriteOffice-help.xml'),
    (Join-Path $resolvedRepoRoot 'Docs\Generated\PSWriteOffice-help.xml'),
    (Join-Path $resolvedRepoRoot 'Artefacts\Unpacked\Modules\PSWriteOffice\en-US\PSWriteOffice-help.xml')
) | Select-Object -Unique

$resolvedHelpPath = $helpCandidates |
    Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } |
    Select-Object -First 1

if ($resolvedHelpPath) {
    Copy-Item -LiteralPath $resolvedHelpPath -Destination $targetHelpPath -Force
    $summary.helpSource = $resolvedHelpPath
    $summary.helpUpdated = $true
    $summary.fallbackUsed = $false
} else {
    Write-Host 'PSWriteOffice help XML not found in synced repo. Keeping checked-in fallback help snapshot.' -ForegroundColor Yellow
}

$manifestCandidates = @(
    (Join-Path $resolvedRepoRoot 'WebsiteArtifacts\apidocs\powershell\PSWriteOffice.psd1'),
    (Join-Path $resolvedRepoRoot 'Artefacts\Unpacked\Modules\PSWriteOffice\PSWriteOffice.psd1'),
    (Join-Path $resolvedRepoRoot 'PSWriteOffice.psd1')
) | Select-Object -Unique

$resolvedManifestPath = $manifestCandidates |
    Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } |
    Select-Object -First 1

if ($resolvedManifestPath) {
    Copy-Item -LiteralPath $resolvedManifestPath -Destination $targetManifestPath -Force
    $summary.manifestSource = $resolvedManifestPath
    $summary.manifestUpdated = $true
    $summary.fallbackUsed = $false
} else {
    Write-Host 'PSWriteOffice module manifest not found in synced repo. Keeping checked-in fallback manifest snapshot.' -ForegroundColor Yellow
}

$commandMetadataCandidates = @(
    (Join-Path $resolvedRepoRoot 'WebsiteArtifacts\apidocs\powershell\command-metadata.json')
) | Select-Object -Unique

$resolvedCommandMetadataPath = $commandMetadataCandidates |
    Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } |
    Select-Object -First 1

if ($resolvedCommandMetadataPath) {
    Copy-Item -LiteralPath $resolvedCommandMetadataPath -Destination $targetCommandMetadataPath -Force
    $summary.commandMetadataSource = $resolvedCommandMetadataPath
    $summary.commandMetadataUpdated = $true
    $summary.fallbackUsed = $false
} else {
    Write-Host 'PSWriteOffice command metadata not found in synced repo. Keeping checked-in fallback command metadata snapshot.' -ForegroundColor Yellow
}

if (-not $SkipExamples) {
    $sourceExamplesPath = Join-Path $resolvedRepoRoot 'Examples'
    if (Test-Path -LiteralPath $sourceExamplesPath -PathType Container) {
        Sync-DirectoryContents -Source $sourceExamplesPath -Destination $targetExamplesPath
        $summary.examplesSource = $sourceExamplesPath
        $summary.examplesUpdated = $true
        $summary.fallbackUsed = $false
    } else {
        Write-Host 'PSWriteOffice examples folder not found in synced repo. Keeping checked-in fallback examples.' -ForegroundColor Yellow
    }
}

[PSCustomObject] $summary
