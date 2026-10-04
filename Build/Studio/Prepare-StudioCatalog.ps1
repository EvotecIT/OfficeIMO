param(
    [Parameter(Mandatory)]
    [string] $AssetRoot,
    [Parameter(Mandatory)]
    [string] $OutputPath
)

$ErrorActionPreference = 'Stop'
powerforge release prepare-catalog --config (Join-Path $PSScriptRoot 'powerforge.windows-release.json') `
    --manifest (Join-Path $AssetRoot 'windows-release-manifest.json') `
    --checksums (Join-Path $AssetRoot 'windows-SHA256SUMS.txt') `
    --asset-root $AssetRoot --out $OutputPath
if ($LASTEXITCODE -ne 0) { throw 'Studio catalog preparation failed.' }
