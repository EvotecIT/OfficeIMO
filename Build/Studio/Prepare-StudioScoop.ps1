<#
.SYNOPSIS
Prepares a Scoop manifest from downloaded Studio release archives.
.DESCRIPTION
PowerForge verifies the release manifest, checksums and portable ZIPs. This does not build, sign, reserve a version or publish.
.PARAMETER AssetRoot
Directory containing the Windows release manifest, checksums, and both portable ZIPs.
.PARAMETER OutputPath
New directory for the generated Scoop manifest.
.EXAMPLE
./Build/Studio/Prepare-StudioScoop.ps1 -AssetRoot ./Artifacts/Studio/Downloads -OutputPath ./Artifacts/Studio/Scoop
#>
param(
    [Parameter(Mandatory)][string] $AssetRoot,
    [Parameter(Mandatory)][string] $OutputPath
)

$ErrorActionPreference = 'Stop'
& powerforge release prepare-scoop --config (Join-Path $PSScriptRoot 'powerforge.windows-release.json') `
    --manifest (Join-Path $AssetRoot 'windows-release-manifest.json') `
    --checksums (Join-Path $AssetRoot 'windows-SHA256SUMS.txt') --asset-root $AssetRoot --out $OutputPath
if ($LASTEXITCODE -ne 0) { throw 'Scoop manifest preparation failed. Inspect the PowerForge error before retrying.' }
