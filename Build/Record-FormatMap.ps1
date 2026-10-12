<#
.SYNOPSIS
    Re-records the format-map marketing clips from the current catalog.

.DESCRIPTION
    Runs two steps in order and stops at the first failure:
      1. Regenerates the catalog data (Website/data/format_map.json) from the conversion catalog and powershell-routes.json.
      2. Renders the selected cuts frame by frame and encodes them with ffmpeg (default recipe: Build/FormatMapMedia/cuts.json).
    The diagram (Build/FormatMapMedia/index.html) reads the generated data directly, so no website build is needed.
    Clips and poster frames land in Artefacts/FormatMapVideos (not tracked). Video encoding needs ffmpeg with the selected codecs;
    PNG stills need only the shared browser. Encoder availability is checked before rendering. Rendering or encoding failures preserve existing clips, posters and stills.
    Failed file replacements restore previous outputs. If recovery is blocked, the error identifies the retained recovery directory.
    Custom cuts must set posterMs within their duration; poster timestamps are checked before frame capture.
    Cut JSON rejects unknown or duplicate properties and missing required fields; configured playback speed is preserved without rounding.

.PARAMETER CutsFile
    Cut JSON file. Default: Build/FormatMapMedia/cuts.json. Use Build/FormatMapMedia/social-cuts.json for platform exports.

.PARAMETER Cut
    Only these cuts (names from the selected cut file). Default: all.

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
    Write PNG stills of these moments (milliseconds) instead of rendering video. Timestamps must fit the selected cuts' durations.

.PARAMETER List
    Print the cuts and exit.

.EXAMPLE
    ./Build/Record-FormatMap.ps1
    ./Build/Record-FormatMap.ps1 -CutsFile Build/FormatMapMedia/social-cuts.json
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
    [switch] $List,
    [ValidateNotNullOrEmpty()]
    [string] $CutsFile
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
    $cutArguments = @()
    if ($PSBoundParameters.ContainsKey('CutsFile')) { $cutArguments += "--cuts=$CutsFile" }
    if ($List) {
        dotnet run --project $recorder -c Release -- @cutArguments --list
        exit $LASTEXITCODE
    }

    if (-not $SkipCatalog) {
        Invoke-Step 'Regenerate catalog data' {
            dotnet run --project (Join-Path $repository 'Build/CompatibilityCatalog/OfficeIMO.CompatibilityCatalog.Tool.csproj') -c Release -f net10.0
        }
    }

    $arguments = @($cutArguments)
    if ($PSBoundParameters.ContainsKey('Still')) {
        if ($Still.Count -eq 0) { throw 'Still requires at least one timestamp.' }
        $arguments += "--still=$($Still -join ',')"
    }
    if ($PSBoundParameters.ContainsKey('Cut')) { $arguments += "--cut=$($Cut -join ',')" }
    if ($Scale -gt 1) { $arguments += "--scale=$Scale" }
    if ($PSBoundParameters.ContainsKey('Fps')) { $arguments += "--fps=$Fps" }
    if ($PSBoundParameters.ContainsKey('Formats')) { $arguments += "--formats=$($Formats -join ',')" }
    if ($Output) { $arguments += "--output=$Output" }
    Invoke-Step 'Render clips' {
        dotnet run --project $recorder -c Release -- @arguments
    }
}
finally {
    Pop-Location
}
