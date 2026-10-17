<#
.SYNOPSIS
Packages a built OfficeIMO Apple engine and its exact dependency notices using PowerForge.
.DESCRIPTION
Build the engine with OfficeIMO.Studio.Native/Build-AppleEngine.ps1 first. This wrapper stages the product's framework
and license inputs; PowerForge owns notice validation, archive creation, metadata and GitHub release publication.
Publish requires the source revision used to build the engine and GITHUB_TOKEN. Existing release assets are never replaced.
PowerForgeModule selects the PSPublishModule assembly built beside the pinned CLI; installed modules are not used.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$EngineDirectory,
    [Parameter(Mandatory)][string]$OutputDirectory,
    [Parameter(Mandatory)][string]$PowerForgeCli,
    [string]$PowerForgeModule,
    [Parameter(Mandatory)][ValidatePattern('^[a-fA-F0-9]{40}$')][string]$SourceCommit,
    [ValidatePattern('^\d+\.\d+\.\d+$')][string]$Version = '1.0.0',
    [bool]$Publish = $false
)
$ErrorActionPreference = 'Stop'
if ($Publish) {
    if ([string]::IsNullOrWhiteSpace($PowerForgeModule)) { throw 'Publish requires the PSPublishModule assembly built from the pinned PowerForge checkout.' }
    $publisher = Import-Module -Name (Resolve-Path -LiteralPath $PowerForgeModule).Path -PassThru -ErrorAction Stop
    $publishRelease = $publisher.ExportedCmdlets['Send-GitHubRelease']
    if ($null -eq $publishRelease) { throw 'The pinned module does not export Send-GitHubRelease.' }
}
$root = (Resolve-Path (Join-Path $PSScriptRoot '../../..')).Path
$EngineDirectory = (Resolve-Path $EngineDirectory).Path
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
$OutputDirectory = (Resolve-Path $OutputDirectory).Path
$payload = Join-Path $OutputDirectory 'payload'
New-Item -ItemType Directory -Path $payload -ErrorAction Stop | Out-Null
Copy-Item (Join-Path $EngineDirectory 'OfficeIMOStudioEngine.xcframework') $payload -Recurse
$licenses = Join-Path $payload 'Licenses'
foreach ($rid in 'osx-arm64', 'osx-x64', 'ios-arm64', 'iossimulator-arm64') {
    & dotnet $PowerForgeCli dotnet notices --project-root $root --deps-json (Join-Path $EngineDirectory "link-inputs/$rid.deps.json") `
        --manifest (Join-Path $PSScriptRoot 'third-party-notices.json') --out (Join-Path $licenses $rid)
    if ($LASTEXITCODE) { throw "PowerForge notice validation failed for $rid." }
}
$items = & dotnet msbuild (Join-Path $PSScriptRoot '../Studio.Licenses.props') -getItem:StudioLicense -nologo | ConvertFrom-Json
if ($LASTEXITCODE) { throw 'Cannot resolve the canonical Studio license resources.' }
foreach ($item in $items.Items.StudioLicense) {
    $destination = Join-Path $payload $item.Link
    New-Item -ItemType Directory -Path (Split-Path $destination) -Force | Out-Null
    Copy-Item -LiteralPath $item.FullPath -Destination $destination
}
$config = Join-Path $PSScriptRoot 'powerforge.engine-package.json'
$tokens = @('--token', "version=$Version", '--token', "sourceCommit=$SourceCommit")
& dotnet $PowerForgeCli dotnet bundle-postprocess --config $config --bundle Studio.Apple.Engine --bundle-root $OutputDirectory --skip-metadata @tokens
if ($LASTEXITCODE) { throw 'PowerForge engine archive failed.' }
$archive = Join-Path $OutputDirectory "OfficeIMOStudioEngine-$Version.zip"
$hash = (Get-FileHash -LiteralPath $archive -Algorithm SHA256).Hash.ToLowerInvariant()
Set-Content -LiteralPath "$archive.sha256" -Value "$hash  $([IO.Path]::GetFileName($archive))" -Encoding utf8NoBOM
& dotnet $PowerForgeCli dotnet bundle-postprocess --config $config --bundle Studio.Apple.Engine --bundle-root $OutputDirectory --skip-archive @tokens --token "archiveSha256=$hash"
if ($LASTEXITCODE) { throw 'PowerForge engine provenance metadata failed.' }
if ($Publish) {
    $release = & $publishRelease -GitHubUsername EvotecIT -GitHubRepositoryName OfficeIMO -GitHubAccessToken $env:GITHUB_TOKEN `
        -TagName "studio-engine-v$Version" -ReleaseName "OfficeIMO Studio Engine $Version" -Commitish $SourceCommit `
        -IsPreRelease $true -ReleaseNotes "Native Apple engine artifact. Source: $SourceCommit. License: MIT; third-party notices accompany the artifact." `
        -AssetFilePaths @($archive, "$archive.sha256", (Join-Path $OutputDirectory 'engine-package.json'))
    if (-not $release.Succeeded) { throw 'PowerForge engine release publication failed.' }
}
Write-Output "SwiftPM checksum: $hash"
Write-Output "Bundle license resources: $licenses"
