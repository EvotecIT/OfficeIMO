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
$inputs = Join-Path $package 'link-inputs'
$stamp = Join-Path $package 'OfficeIMOStudioEngine.source-key'

function Invoke-Git([string[]]$Arguments, [string[]]$InputObject) {
    $output = if ($PSBoundParameters.ContainsKey('InputObject')) { $InputObject | git -C $root @Arguments } else { git -C $root @Arguments }
    if ($LASTEXITCODE -ne 0) { throw "git $($Arguments[0]) failed while reading the engine sources." }
    return $output
}

function Get-SourceKey {
    $spec = @('--', '.', ':(exclude)Apple', ':(exclude)Website', ':(exclude)Docs', ':(exclude).github', ':(exclude)*.md')
    $tracked = Invoke-Git (@('ls-files', '-s') + $spec)
    $changes = Invoke-Git (@('diff', '--no-ext-diff', '--binary') + $spec)
    $untracked = @(Invoke-Git (@('ls-files', '--others', '--exclude-standard') + $spec))
    $untrackedHashes = if ($untracked.Count -gt 0) { Invoke-Git @('hash-object', '--stdin-paths') $untracked } else { @() }
    # The engine statically links the NativeAOT runtime, so an SDK update needs a new build even when the sources match.
    Push-Location $root
    try { $sdk = dotnet --version } finally { Pop-Location }
    if ($LASTEXITCODE -ne 0) { throw 'Cannot resolve the .NET SDK for this checkout.' }
    $text = (@($tracked) + @($changes) + $untracked + @($untrackedHashes) + ($RuntimeIdentifier | Sort-Object) + "sdk $sdk") -join "`n"
    $bytes = [System.Security.Cryptography.SHA256]::HashData([System.Text.Encoding]::UTF8.GetBytes($text))
    return [System.Convert]::ToHexString($bytes).ToLowerInvariant()
}

$key = Get-SourceKey
if ($PrintSourceKey) { return $key }

if (-not $Force -and (Test-Path $framework) -and (Test-Path $stamp) -and (Get-Content -Raw $stamp).Trim() -eq $key) {
    Write-Host "The engine is up to date with this checkout ($($key.Substring(0, 12)))."
    return
}

# Installs the framework and its link inputs (dependency manifests for the third-party notices) from a build or cache
# folder. The stamp is written last, so an interrupted install is never taken as up to date.
function Install-Engine([string]$Source) {
    if (Test-Path $stamp) { Remove-Item -Force $stamp }
    foreach ($item in $framework, $inputs) { if (Test-Path $item) { Remove-Item -Recurse -Force $item } }
    Copy-Item -Recurse (Join-Path $Source 'OfficeIMOStudioEngine.xcframework') $framework
    Copy-Item -Recurse (Join-Path $Source 'link-inputs') $inputs
    Set-Content -LiteralPath $stamp -Value $key -NoNewline
}

$cached = if ($CacheDirectory) { Join-Path $CacheDirectory $key } else { $null }
if (-not $Force -and $cached -and (Test-Path (Join-Path $cached 'OfficeIMOStudioEngine.xcframework')) -and (Test-Path (Join-Path $cached 'link-inputs'))) {
    Install-Engine $cached
    Write-Host "Installed the cached engine for this checkout ($($key.Substring(0, 12)))."
    return
}

if (-not $WorkDirectory) { $WorkDirectory = Join-Path $package '.build/engine' }
& (Join-Path $root 'OfficeIMO.Studio.Native/Build-AppleEngine.ps1') -OutputDirectory $WorkDirectory -RuntimeIdentifier $RuntimeIdentifier
Install-Engine $WorkDirectory

if ($CacheDirectory) {
    # Copied into a staging folder and renamed once complete, so an interrupted copy never looks like a cache hit.
    $entry = Join-Path $CacheDirectory $key
    $staging = "$entry.tmp"
    if (Test-Path $staging) { Remove-Item -Recurse -Force $staging }
    New-Item -ItemType Directory -Force $staging | Out-Null
    Copy-Item -Recurse $framework (Join-Path $staging 'OfficeIMOStudioEngine.xcframework')
    Copy-Item -Recurse $inputs (Join-Path $staging 'link-inputs')
    if (Test-Path $entry) { Remove-Item -Recurse -Force $entry }
    Move-Item $staging $entry
    Get-ChildItem -Directory $CacheDirectory | Where-Object Name -NotLike '*.tmp' |
        Sort-Object LastWriteTimeUtc -Descending | Select-Object -Skip 3 |
        ForEach-Object { Remove-Item -Recurse -Force $_.FullName }
}
Write-Host "Built the engine for this checkout ($($key.Substring(0, 12)))."
