#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Build and serve the OfficeIMO website.

.PARAMETER Serve
    Build the full pipeline and start the development server.

.PARAMETER Fast
    Run the pipeline with `--fast` (recommended for local iteration).
    When `-Serve` is used, fast mode is enabled by default unless `-NoFast` is set.

.PARAMETER Dev
    Run the pipeline with `--dev` (implies fast mode and enables pipeline mode 'dev').
    When `-Serve` is used, dev mode is enabled by default unless `-NoDev` is set.

.PARAMETER NoFast
    Disable fast mode when `-Serve` is used.

.PARAMETER NoDev
    Disable dev mode when `-Serve` is used.

.PARAMETER Only
    Run only the specified pipeline tasks (comma/semicolon separated), for example: build,verify.

.PARAMETER Skip
    Skip the specified pipeline tasks (comma/semicolon separated), for example: optimize,audit.

.PARAMETER Watch
    Run the pipeline in watch mode (rebuild on file changes).
    When combined with `-Serve`, starts the static server and keeps rebuilding in the foreground.

.PARAMETER Port
    Server port (default: 8080).

.PARAMETER PowerForgeRoot
    Root folder of PSPublishModule (overrides $env:POWERFORGE_ROOT).

.PARAMETER CI
    Force pipeline CI mode locally (equivalent to setting CI=true).

.PARAMETER PSWriteOfficeRoot
    Optional PSWriteOffice repo root used to refresh the checked-in PowerShell API snapshot before building.

.EXAMPLE
    ./build.ps1
    ./build.ps1 -CI
    ./build.ps1 -Serve
    ./build.ps1 -Serve -Port 3000
#>

param(
    [switch]$Serve,
    [switch]$Watch,
    [switch]$Dev,
    [switch]$Fast,
    [switch]$NoDev,
    [switch]$NoFast,
    [int]$Port = 8080,
    [string[]]$Only = @(),
    [string[]]$Skip = @(),
    [switch]$CI,
    [switch]$SkipBuildTool,
    [string]$PowerForgeRoot = $env:POWERFORGE_ROOT,
    [string]$PSWriteOfficeRoot = ''
)

$ErrorActionPreference = 'Stop'
Push-Location $PSScriptRoot

function Resolve-PowerForgeRoot {
    param(
        [string]$RequestedRoot
    )

    $candidates = @()
    if (-not [string]::IsNullOrWhiteSpace($RequestedRoot)) {
        $candidates += $RequestedRoot
    }
    if (-not [string]::IsNullOrWhiteSpace($env:POWERFORGE_ROOT)) {
        $candidates += $env:POWERFORGE_ROOT
    }

    # Walk up and probe for a sibling PSPublishModule folder (works for common mono-root layouts).
    $here = (Resolve-Path -LiteralPath $PSScriptRoot).Path
    $cursor = $here
    for ($i = 0; $i -lt 8; $i++) {
        $parent = Split-Path -Parent $cursor
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $cursor) { break }
        $candidates += (Join-Path $parent 'PSPublishModule')
        $cursor = $parent
    }

    # Maintainer defaults + WSL-style mounts.
    if ($IsWindows) {
        $candidates += 'C:\Support\GitHub\PSPublishModule'
    }
    $candidates += '/mnt/c/Support/GitHub/PSPublishModule'

    foreach ($candidate in ($candidates | Select-Object -Unique)) {
        if ([string]::IsNullOrWhiteSpace($candidate)) { continue }
        try {
            $root = (Resolve-Path -LiteralPath $candidate -ErrorAction Stop).Path
        } catch {
            continue
        }

        $cliProject = Join-Path $root 'PowerForge.Web.Cli\PowerForge.Web.Cli.csproj'
        if (Test-Path -LiteralPath $cliProject -PathType Leaf) {
            return $root
        }
    }

    return $null
}

# Resolve PowerForge.Web.Cli executable (prefer local fresh build)
$PowerForge = $null
$PowerForgeArgs = @()
$PowerForgeRoot = Resolve-PowerForgeRoot -RequestedRoot $PowerForgeRoot
if (-not [string]::IsNullOrWhiteSpace($PowerForgeRoot)) {
    $PowerForgeCliProject = Join-Path $PowerForgeRoot 'PowerForge.Web.Cli\PowerForge.Web.Cli.csproj'

    $tfms = @('net10.0', 'net8.0')
    foreach ($tfm in $tfms) {
        if (-not $PowerForgeReleaseExe) { $PowerForgeReleaseExe = Join-Path $PowerForgeRoot "PowerForge.Web.Cli\\bin\\Release\\$tfm\\PowerForge.Web.Cli.exe" }
        if (-not $PowerForgeReleaseAppHost) { $PowerForgeReleaseAppHost = Join-Path $PowerForgeRoot "PowerForge.Web.Cli\\bin\\Release\\$tfm\\PowerForge.Web.Cli" }
        if (-not $PowerForgeReleaseDll) { $PowerForgeReleaseDll = Join-Path $PowerForgeRoot "PowerForge.Web.Cli\\bin\\Release\\$tfm\\PowerForge.Web.Cli.dll" }

        if (-not $PowerForgeDebugExe) { $PowerForgeDebugExe = Join-Path $PowerForgeRoot "PowerForge.Web.Cli\\bin\\Debug\\$tfm\\PowerForge.Web.Cli.exe" }
        if (-not $PowerForgeDebugAppHost) { $PowerForgeDebugAppHost = Join-Path $PowerForgeRoot "PowerForge.Web.Cli\\bin\\Debug\\$tfm\\PowerForge.Web.Cli" }
        if (-not $PowerForgeDebugDll) { $PowerForgeDebugDll = Join-Path $PowerForgeRoot "PowerForge.Web.Cli\\bin\\Debug\\$tfm\\PowerForge.Web.Cli.dll" }
    }
}

if (-not $SkipBuildTool -and $PowerForgeCliProject -and (Test-Path $PowerForgeCliProject)) {
    Write-Host "Building PowerForge.Web.Cli..." -ForegroundColor Cyan
    dotnet build $PowerForgeCliProject -c Release | Out-Host
    if ($LASTEXITCODE -ne 0) { throw "PowerForge.Web.Cli build failed (exit code $LASTEXITCODE)" }
}

if ($PowerForgeReleaseExe -and (Test-Path $PowerForgeReleaseExe)) {
    $PowerForge = $PowerForgeReleaseExe
} elseif ($PowerForgeReleaseAppHost -and (Test-Path $PowerForgeReleaseAppHost)) {
    $PowerForge = $PowerForgeReleaseAppHost
} elseif ($PowerForgeReleaseDll -and (Test-Path $PowerForgeReleaseDll)) {
    $PowerForge = 'dotnet'
    $PowerForgeArgs = @($PowerForgeReleaseDll)
} elseif ($PowerForgeDebugExe -and (Test-Path $PowerForgeDebugExe)) {
    $PowerForge = $PowerForgeDebugExe
} elseif ($PowerForgeDebugAppHost -and (Test-Path $PowerForgeDebugAppHost)) {
    $PowerForge = $PowerForgeDebugAppHost
} elseif ($PowerForgeDebugDll -and (Test-Path $PowerForgeDebugDll)) {
    $PowerForge = 'dotnet'
    $PowerForgeArgs = @($PowerForgeDebugDll)
} else {
    $PowerForge = Get-Command powerforge-web -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source
}

if (-not $PowerForge) {
    Write-Error 'PowerForge.Web.Cli not found. Install the global tool (powerforge-web), set POWERFORGE_ROOT (path to PSPublishModule), or pass -PowerForgeRoot.'
    exit 1
}

Write-Host "Using: $PowerForge $($PowerForgeArgs -join ' ')" -ForegroundColor DarkGray

function Assert-SiteOutput {
    param(
        [Parameter(Mandatory)]
        [string]$SiteRoot
    )

    $notFoundPage = Join-Path $SiteRoot '404.html'
    if (-not (Test-Path -LiteralPath $notFoundPage -PathType Leaf)) {
        throw "Build validation failed: expected '$notFoundPage' for GitHub Pages 404 handling."
    }
}

try {
    if (-not [string]::IsNullOrWhiteSpace($PSWriteOfficeRoot)) {
        $syncScript = Join-Path $PSScriptRoot 'scripts\Sync-PSWriteOfficeApiDocs.ps1'
        if (-not (Test-Path -LiteralPath $syncScript -PathType Leaf)) {
            throw "PSWriteOffice API sync script not found: $syncScript"
        }

        Write-Host 'Refreshing PSWriteOffice API snapshot...' -ForegroundColor Cyan
        & $syncScript -SiteRoot $PSScriptRoot -PSWriteOfficeRoot $PSWriteOfficeRoot
        if (-not $?) { throw 'PSWriteOffice API sync failed.' }
    }

    $UseDev = ($Dev -or ($Serve -and -not $NoDev))
    $UseFast = ($Fast -or ($Serve -and -not $NoFast))
    $IsCI = $CI -or
            ($env:CI -and $env:CI.ToString().ToLowerInvariant() -eq 'true') -or
            ($env:GITHUB_ACTIONS -and $env:GITHUB_ACTIONS.ToString().ToLowerInvariant() -eq 'true') -or
            ($env:TF_BUILD -and $env:TF_BUILD.ToString().ToLowerInvariant() -eq 'true')

    $pipelineArgsBase = @('pipeline', '--config', 'pipeline.json', '--profile')
    if ($UseDev) {
        $pipelineArgsBase += '--dev'
    } elseif ($UseFast) {
        $pipelineArgsBase += '--fast'
    }
    if ($IsCI -and -not $UseDev) {
        $pipelineArgsBase += @('--mode', 'ci')
    }
    if ($Only -and $Only.Count -gt 0) {
        $pipelineArgsBase += @('--only', ($Only -join ','))
    }
    if ($Skip -and $Skip.Count -gt 0) {
        $pipelineArgsBase += @('--skip', ($Skip -join ','))
    }

    $pipelineArgs = @($pipelineArgsBase)
    if ($Watch) {
        $pipelineArgs += '--watch'
    }

    $modeLabel = $UseDev ? 'dev' : ($UseFast ? 'fast' : ($IsCI ? 'ci' : 'default'))
    Write-Host "Pipeline mode: $modeLabel" -ForegroundColor DarkGray

    if ($Serve) {
        Write-Host 'Building website...' -ForegroundColor Cyan
        if ($Watch) {
            & $PowerForge @PowerForgeArgs @pipelineArgsBase
        } else {
            & $PowerForge @PowerForgeArgs @pipelineArgs
        }
        if ($LASTEXITCODE -ne 0) { throw "Build failed (exit code $LASTEXITCODE)" }
        Assert-SiteOutput -SiteRoot (Join-Path $PSScriptRoot '_site')
        Write-Host "Starting dev server on http://localhost:$Port ..." -ForegroundColor Cyan

        $serveArgs = @($PowerForgeArgs + @('serve', '--path', '_site', '--port', $Port))
        $serveProcess = Start-Process -FilePath $PowerForge -ArgumentList $serveArgs -NoNewWindow -PassThru
        try {
            if ($Watch) {
                Write-Host 'Watching for changes...' -ForegroundColor Cyan
                & $PowerForge @PowerForgeArgs @pipelineArgs
            } else {
                Wait-Process -Id $serveProcess.Id
            }
        } finally {
            if ($serveProcess -and -not $serveProcess.HasExited) {
                Stop-Process -Id $serveProcess.Id -Force -ErrorAction SilentlyContinue
            }
        }
    } else {
        Write-Host 'Building website...' -ForegroundColor Cyan
        & $PowerForge @PowerForgeArgs @pipelineArgs
        if ($LASTEXITCODE -ne 0) { throw "Build failed (exit code $LASTEXITCODE)" }
        Assert-SiteOutput -SiteRoot (Join-Path $PSScriptRoot '_site')
        Write-Host 'Build complete -> _site/' -ForegroundColor Green
    }
} finally {
    Pop-Location
}
