<#
.SYNOPSIS
Builds the Studio engine from this checkout into the local engine package, when its sources changed.

.DESCRIPTION
The Mac, iPad and iPhone apps link Apple/OfficeIMOStudioEngine/OfficeIMOStudioEngine.xcframework, which is built from
the same checkout and never committed, so a merged engine change reaches the next app build without publishing
anything. This script builds it with OfficeIMO.Studio.Native/Build-AppleEngine.ps1 and records a key for the engine
sources and runtimes it was built from. A later run with the same key does nothing.

Run it before building the apps in Xcode, in CI and before an App Store archive. The key covers every tracked file
outside Apple/, Website/, Docs/, .github/ and Markdown files, including uncommitted and untracked changes.

With -CacheDirectory (or OFFICEIMO_STUDIO_ENGINE_CACHE), builds are also kept there by key, so a runner or a second
checkout reuses them. The newest three are kept.

.EXAMPLE
./Build/Studio/Apple/Build-StudioEngine.ps1

.EXAMPLE
./Build/Studio/Apple/Build-StudioEngine.ps1 -RuntimeIdentifier osx-arm64, iossimulator-arm64
#>
[CmdletBinding()]
param(
    [ValidateSet('osx-arm64', 'osx-x64', 'ios-arm64', 'iossimulator-arm64')]
    [string[]]$RuntimeIdentifier = @('osx-arm64', 'osx-x64', 'ios-arm64', 'iossimulator-arm64'),
    # Build output and intermediates. Defaults to the engine package's ignored .build folder.
    [string]$WorkDirectory,
    [string]$CacheDirectory = $env:OFFICEIMO_STUDIO_ENGINE_CACHE,
    # Prints the source key and exits.
    [switch]$PrintSourceKey,
    [switch]$Force
)
$ErrorActionPreference = 'Stop'

$root = (Resolve-Path (Join-Path $PSScriptRoot '../../..')).Path
$package = Join-Path $root 'Apple/OfficeIMOStudioEngine'
$framework = Join-Path $package 'OfficeIMOStudioEngine.xcframework'
$stamp = Join-Path $package 'OfficeIMOStudioEngine.source-key'

function Get-SourceKey {
    $spec = @('--', '.', ':(exclude)Apple', ':(exclude)Website', ':(exclude)Docs', ':(exclude).github', ':(exclude)*.md')
    $tracked = git -C $root ls-files -s @spec
    if ($LASTEXITCODE -ne 0) { throw 'Cannot list the engine sources with git.' }
    $changes = git -C $root diff --no-ext-diff --binary @spec
    $untracked = @(git -C $root ls-files --others --exclude-standard @spec)
    $untrackedHashes = if ($untracked.Count -gt 0) { $untracked | git -C $root hash-object --stdin-paths } else { @() }
    $text = (@($tracked) + @($changes) + $untracked + @($untrackedHashes) + ($RuntimeIdentifier | Sort-Object)) -join "`n"
    $bytes = [System.Security.Cryptography.SHA256]::HashData([System.Text.Encoding]::UTF8.GetBytes($text))
    return [System.Convert]::ToHexString($bytes).ToLowerInvariant()
}

$key = Get-SourceKey
if ($PrintSourceKey) { return $key }

if (-not $Force -and (Test-Path $framework) -and (Test-Path $stamp) -and (Get-Content -Raw $stamp).Trim() -eq $key) {
    Write-Host "The engine is up to date with this checkout ($($key.Substring(0, 12)))."
    return
}

function Install-Engine([string]$Source) {
    if (Test-Path $framework) { Remove-Item -Recurse -Force $framework }
    Copy-Item -Recurse $Source $framework
    Set-Content -LiteralPath $stamp -Value $key -NoNewline
}

$cached = if ($CacheDirectory) { Join-Path $CacheDirectory $key 'OfficeIMOStudioEngine.xcframework' } else { $null }
if (-not $Force -and $cached -and (Test-Path $cached)) {
    Install-Engine $cached
    Write-Host "Installed the cached engine for this checkout ($($key.Substring(0, 12)))."
    return
}

if (-not $WorkDirectory) { $WorkDirectory = Join-Path $package '.build/engine' }
& (Join-Path $root 'OfficeIMO.Studio.Native/Build-AppleEngine.ps1') -OutputDirectory $WorkDirectory -RuntimeIdentifier $RuntimeIdentifier
Install-Engine (Join-Path $WorkDirectory 'OfficeIMOStudioEngine.xcframework')

if ($CacheDirectory) {
    $entry = Join-Path $CacheDirectory $key
    if (Test-Path $entry) { Remove-Item -Recurse -Force $entry }
    New-Item -ItemType Directory -Force $entry | Out-Null
    Copy-Item -Recurse $framework (Join-Path $entry 'OfficeIMOStudioEngine.xcframework')
    Get-ChildItem -Directory $CacheDirectory | Sort-Object LastWriteTimeUtc -Descending | Select-Object -Skip 3 |
        ForEach-Object { Remove-Item -Recurse -Force $_.FullName }
}
Write-Host "Built the engine for this checkout ($($key.Substring(0, 12)))."
