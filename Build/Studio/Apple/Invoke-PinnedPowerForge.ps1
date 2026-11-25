[CmdletBinding(PositionalBinding = $false)]
param(
    [string] $PSPublishModuleRoot = $env:PSPUBLISHMODULE_ROOT,

    [ValidateSet('Apple', 'Web')]
    [string] $Tool = 'Apple',

    [Parameter(Mandatory, Position = 0, ValueFromRemainingArguments)]
    [string[]] $ArgumentList
)

# Runs PowerForge from a PSPublishModule checkout at the reviewed commit below.
# The Studio Apple workflows pin the same commit.
$ErrorActionPreference = 'Stop'
$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '../../..')).Path
$expectedPowerForgeCommit = '2664d018b08c69a6eed5622135b30cc6c9d2428c'

if ([string]::IsNullOrWhiteSpace($PSPublishModuleRoot)) {
    $candidates = @()
    if (-not [string]::IsNullOrWhiteSpace($env:EVOTEC_GITHUB_ROOT)) {
        $candidates += Join-Path $env:EVOTEC_GITHUB_ROOT 'PSPublishModule'
    }
    $candidates += Join-Path (Split-Path -Parent $repoRoot) 'PSPublishModule'
    $candidates += Join-Path (Split-Path -Parent (Split-Path -Parent $repoRoot)) 'PSPublishModule'
    $PSPublishModuleRoot = $candidates |
        Where-Object { Test-Path -LiteralPath $_ -PathType Container } |
        Select-Object -First 1
}

if ([string]::IsNullOrWhiteSpace($PSPublishModuleRoot)) {
    throw 'PSPublishModuleRoot was not supplied and no sibling PSPublishModule checkout was found.'
}

$helper = Join-Path $PSPublishModuleRoot 'scripts/Invoke-PinnedPowerForge.ps1'
if (-not (Test-Path -LiteralPath $helper -PathType Leaf)) {
    throw "Pinned PowerForge helper was not found: $helper"
}

& $helper `
    -ExpectedCommit $expectedPowerForgeCommit `
    -Tool $Tool `
    -ConsumerRoot $repoRoot `
    -ExpectedConsumerRepository 'EvotecIT/OfficeIMO' `
    -ExpectedConsumerBranch master `
    -ArgumentList $ArgumentList

exit $LASTEXITCODE
