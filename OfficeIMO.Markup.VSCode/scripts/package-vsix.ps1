[CmdletBinding()]
param(
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Release',

    [string]$Framework = 'net8.0',

    [string[]]$RuntimeIdentifiers = @('win-x64', 'win-arm64', 'linux-x64', 'linux-arm64', 'osx-x64', 'osx-arm64'),

    [string]$OutputDirectory = 'dist',

    [switch]$SkipNpmCi,

    [switch]$SkipRestore,

    [switch]$PublishMarketplace,

    [switch]$PreRelease,

    [string]$VscePat = $env:VSCE_PAT
)

$ErrorActionPreference = 'Stop'

$packageScript = Join-Path $PSScriptRoot 'package-vsix.cjs'
if (-not (Test-Path -LiteralPath $packageScript)) {
    throw "Package script not found at '$packageScript'."
}

Get-Command node -ErrorAction Stop | Out-Null

$arguments = @(
    $packageScript,
    '--configuration', $Configuration,
    '--framework', $Framework,
    '--runtime-identifiers', ($RuntimeIdentifiers -join ','),
    '--output-directory', $OutputDirectory
)

if ($SkipNpmCi) {
    $arguments += '--skip-npm-ci'
}
if ($SkipRestore) {
    $arguments += '--skip-restore'
}
if ($PublishMarketplace) {
    $arguments += '--publish-marketplace'
}
if ($PreRelease) {
    $arguments += '--pre-release'
}

$previousVscePat = $env:VSCE_PAT
try {
    if (-not [string]::IsNullOrWhiteSpace($VscePat)) {
        $env:VSCE_PAT = $VscePat
    }

    & node @arguments
    if ($LASTEXITCODE -ne 0) {
        throw "'node $($arguments -join ' ')' failed with exit code $LASTEXITCODE."
    }
} finally {
    $env:VSCE_PAT = $previousVscePat
}
