<#
.SYNOPSIS
    Re-records the format-map marketing clips from the current catalog.

.DESCRIPTION
    Runs two steps in order and stops at the first failure:
      1. Regenerates the catalog data (Website/data/format_map.json) from the conversion catalog and powershell-routes.json.
      2. Renders every cut in Build/FormatMapMedia/cuts.json frame by frame and encodes it with ffmpeg.
    The diagram (Build/FormatMapMedia/index.html) reads the generated data directly, so no website build is needed.
    Clips and poster frames land in Artefacts/FormatMapVideos (not tracked). Needs ffmpeg and the shared Playwright browsers.

.PARAMETER Cut
    Only these cuts (names from cuts.json). Default: all.

.PARAMETER Scale
    2 renders 1920x1080 as true 3840x2160 (the diagram is vector, so text and lines stay sharp; about three times slower).

.PARAMETER Fps
    Frames per second of the render (10 to 120). Default: the cut's own, 30.

.PARAMETER Formats
    Replaces each cut's output formats: mp4, h265, av1, webm, mov, gif, webp, apng.

.PARAMETER Output
    Output folder. Default: Artefacts/FormatMapVideos.

.PARAMETER SkipCatalog
    Reuse the current Website/data/format_map.json (the catalog build takes minutes; skip it when only the design or the cuts changed).

.PARAMETER Still
    Write PNG stills of these moments (milliseconds) instead of rendering video. Useful when changing the design.

.PARAMETER List
    Print the cuts and exit.

.EXAMPLE
    ./Build/Record-FormatMap.ps1
    ./Build/Record-FormatMap.ps1 -Cut media-9x16 -SkipCatalog
    ./Build/Record-FormatMap.ps1 -Cut media-16x9 -Scale 2 -SkipCatalog
    ./Build/Record-FormatMap.ps1 -Cut media-9x16 -Still 2000,8000 -SkipCatalog
    ./Build/Record-FormatMap.ps1 -Formats mp4,h265,gif,webp -SkipCatalog
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
    [ValidateRange(0, [int]::MaxValue)]
    [int[]] $Still,
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
    }

    $arguments = @()
    if ($PSBoundParameters.ContainsKey('Still')) {
        if ($Still.Count -eq 0) { throw 'Still requires at least one timestamp.' }
        $arguments += "--still=$($Still -join ',')"
    }
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
