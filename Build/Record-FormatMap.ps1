<#
.SYNOPSIS
    Re-records the format-map marketing clips from the current catalog.

.DESCRIPTION
    Runs the three steps in order and stops at the first failure:
      1. Regenerates the catalog data (Website/data/format_map.json) from the conversion catalog and powershell-routes.json.
      2. Builds the website pages and styles the stage needs (/format-map/).
      3. Renders every cut in Build/FormatMapRecorder/cuts.json frame by frame and encodes it with ffmpeg.
    Clips and poster frames land in Artefacts/FormatMapVideos (not tracked). Needs ffmpeg and the shared Playwright browsers.

.PARAMETER Cut
    Only these cuts (names from cuts.json). Default: all.

.PARAMETER Scale
    2 renders 1920x1080 layouts as true 3840x2160 (about five times slower than 1x).

.PARAMETER Fps
    Frames per second of the render (10 to 120). Default: the cut's own, 30.

.PARAMETER Formats
    Replaces each cut's output formats: mp4, h265, av1, webm, mov, gif, webp, apng.

.PARAMETER Output
    Output folder. Default: Artefacts/FormatMapVideos.

.PARAMETER SkipCatalog
    Reuse the current Website/data/format_map.json (the catalog build takes minutes; skip it when only styles or cuts changed).

.PARAMETER SkipSite
    Reuse the built Website/_site.

.PARAMETER List
    Print the cuts and exit.

.EXAMPLE
    ./Build/Record-FormatMap.ps1
    ./Build/Record-FormatMap.ps1 -Cut powershell-9x16 -SkipCatalog
    ./Build/Record-FormatMap.ps1 -Cut full-16x9 -Scale 2 -SkipCatalog -SkipSite
#>
[CmdletBinding()]
param(
    [string[]] $Cut,
    [ValidateRange(1, 4)]
    [int] $Scale = 1,
    [ValidateRange(10, 120)]
    [int] $Fps,
    [string[]] $Formats,
    [string] $Output,
    [switch] $SkipCatalog,
    [switch] $SkipSite,
    [switch] $List
)

$ErrorActionPreference = 'Stop'
$repository = Split-Path -Parent $PSScriptRoot
$recorder = Join-Path $repository 'Build/FormatMapRecorder/OfficeIMO.FormatMapRecorder.csproj'

function Invoke-Step {
    param([string] $Name, [scriptblock] $Action)
    Write-Host "== $Name" -ForegroundColor Cyan
    & $Action
    if ($LASTEXITCODE -ne 0) { throw "$Name failed (exit code $LASTEXITCODE)." }
}

# The catalog tool writes LF files, so git lists every generated file as modified on a CRLF checkout even when nothing changed.
# Restore the ones whose only difference is line endings; real changes (format_map.json, a changed catalog) are left alone.
function Restore-LineEndingOnlyChanges {
    $changed = @(git diff --name-only 2>$null)
    $restored = 0
    foreach ($file in $changed) {
        git diff --ignore-cr-at-eol --quiet -- $file 2>$null
        if ($LASTEXITCODE -eq 0) { git checkout -- $file 2>$null; $restored++ }
    }
    $global:LASTEXITCODE = 0
    if ($restored -gt 0) { Write-Host "Restored $restored generated files that differed only in line endings." }
}

Push-Location $repository
try {
    if ($List) {
        dotnet run --project $recorder -c Release -- --list
        exit $LASTEXITCODE
    }

    if (-not $SkipCatalog) {
        Invoke-Step 'Regenerate catalog data' {
            dotnet run --project (Join-Path $repository 'Build/CompatibilityCatalog/OfficeIMO.CompatibilityCatalog.Tool.csproj') -c Release -f net10.0
        }
        Restore-LineEndingOnlyChanges
    }

    if (-not $SkipSite) {
        Invoke-Step 'Build website' {
            Push-Location (Join-Path $repository 'Website')
            try { & pwsh -NoProfile -File ./build.ps1 -Dev -Only 'build-site,deploy-theme-css' } finally { Pop-Location }
        }
    }

    $arguments = @()
    if ($Cut) { $arguments += "--cut=$($Cut -join ',')" }
    if ($Scale -gt 1) { $arguments += "--scale=$Scale" }
    if ($PSBoundParameters.ContainsKey('Fps')) { $arguments += "--fps=$Fps" }
    if ($Formats) { $arguments += "--formats=$($Formats -join ',')" }
    if ($Output) { $arguments += "--output=$Output" }
    Invoke-Step 'Render clips' {
        dotnet run --project $recorder -c Release -- @arguments
    }
}
finally {
    Pop-Location
}
