<#
.SYNOPSIS
Loads the PowerForge source revision pinned for library-comparison evidence.
.DESCRIPTION
Uses a dedicated source cache, builds the unsigned PowerShell module, and imports
it. An explicit source checkout is useful for development; its HEAD must match
the pin. PlanOnly returns the pin without fetching, building, or importing.
.PARAMETER PowerForgeRoot
An existing checkout at the pinned revision. Shared checkouts are never changed.
.PARAMETER PlanOnly
Returns the source identity without executing setup.
.EXAMPLE
./Build/Initialize-BenchmarkPowerForge.ps1
.OUTPUTS
The repository, revision and module path used for benchmark evidence tooling.
#>
param(
    [string] $PowerForgeRoot,
    [switch] $PlanOnly
)

$ErrorActionPreference = 'Stop'
$pin = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'benchmark-powerforge.json') -Raw | ConvertFrom-Json
if ($pin.ref -notmatch '^[0-9a-f]{40}$') {
    throw 'The benchmark PowerForge pin must identify an exact Git commit.'
}
if ($PlanOnly) {
    return $pin
}
if ([string]::IsNullOrWhiteSpace($PowerForgeRoot)) {
    $cacheRoot = if (-not [string]::IsNullOrWhiteSpace($env:EVOTEC_SCRATCH_ROOT)) {
        if (-not (Test-Path -LiteralPath $env:EVOTEC_SCRATCH_ROOT -PathType Container)) {
            throw 'The configured EVOTEC_SCRATCH_ROOT is unavailable.'
        }
        $env:EVOTEC_SCRATCH_ROOT
    } else {
        [Environment]::GetFolderPath('LocalApplicationData')
    }
    $PowerForgeRoot = Join-Path $cacheRoot "PowerForge/benchmark-sources/$($pin.ref)"
    if (-not (Test-Path -LiteralPath $PowerForgeRoot)) {
        New-Item -ItemType Directory -Path $PowerForgeRoot -Force | Out-Null
        & git -C $PowerForgeRoot init --quiet
        if ($LASTEXITCODE -ne 0) { throw 'Could not initialize the pinned PowerForge source cache.' }
        & git -C $PowerForgeRoot remote add origin $pin.repository
        if ($LASTEXITCODE -ne 0) { throw 'Could not configure the pinned PowerForge source remote.' }
        & git -C $PowerForgeRoot fetch --depth 1 origin $pin.ref
        if ($LASTEXITCODE -ne 0) { throw 'Could not fetch the pinned PowerForge source revision.' }
        & git -C $PowerForgeRoot checkout --detach --quiet FETCH_HEAD
        if ($LASTEXITCODE -ne 0) { throw 'Could not check out the pinned PowerForge source revision.' }
    }
}
$head = & git -C $PowerForgeRoot rev-parse HEAD
if ($LASTEXITCODE -ne 0 -or $head -ne $pin.ref) {
    throw "PowerForgeRoot must be at the pinned commit $($pin.ref)."
}
$dirty = @(& git -C $PowerForgeRoot status --porcelain --untracked-files=normal)
if ($LASTEXITCODE -ne 0 -or $dirty.Count -gt 0) {
    throw 'The pinned PowerForge source checkout must be clean.'
}
$modulePath = Join-Path $PowerForgeRoot "PSPublishModule/bin/Release/$($pin.framework)/PSPublishModule.dll"
$previousNodeReuse = $env:MSBUILDDISABLENODEREUSE
$previousBuildServer = $env:DOTNET_CLI_USE_MSBUILD_SERVER
try {
    $env:MSBUILDDISABLENODEREUSE = '1'
    $env:DOTNET_CLI_USE_MSBUILD_SERVER = '0'
    Push-Location -LiteralPath $PowerForgeRoot
    try {
        & dotnet build PSPublishModule/PSPublishModule.csproj -c Release -f $pin.framework -m:2 -nodeReuse:false --nologo --verbosity quiet | Out-Host
        if ($LASTEXITCODE -ne 0) { throw 'The pinned PowerForge module build failed.' }
    } finally {
        Pop-Location
    }
    Import-Module $modulePath -Force -Global
} finally {
    $env:MSBUILDDISABLENODEREUSE = $previousNodeReuse
    $env:DOTNET_CLI_USE_MSBUILD_SERVER = $previousBuildServer
}
[pscustomobject]@{
    Repository = $pin.repository
    Ref = $head
    ModulePath = $modulePath
}
