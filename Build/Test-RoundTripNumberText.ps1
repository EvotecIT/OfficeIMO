<#
.SYNOPSIS
Checks the round-trip number text written on .NET Framework against both runtimes' parsers.

.DESCRIPTION
Builds Build/RoundTripNumberText for net472 and net10.0. The net472 run writes the text of the same seeded doubles
and floats and fails when .NET Framework's own parser reads any of them back as another value. The net10.0 run
reads those texts and fails when a correct parser does; it also reports how many equal .NET 10's "R" and how many
are longer because .NET Framework's parser misreads the shortest text. Requires Windows for the net472 build. The
check is opt-in: it takes a minute or two for the default two million values.

.PARAMETER Count
How many doubles and floats to write.

.PARAMETER OutputDirectory
Where the builds and the text files go; a new folder under the temporary directory by default, removed afterwards.
#>
[CmdletBinding()]
param(
    [ValidateRange(1, [int]::MaxValue)]
    [int] $Count = 2000000,
    [string] $OutputDirectory
)

$ErrorActionPreference = 'Stop'
$project = Join-Path $PSScriptRoot 'RoundTripNumberText/OfficeIMO.RoundTripNumberTextProbe.csproj'
$removeOutput = -not $OutputDirectory
if (-not $OutputDirectory) {
    $OutputDirectory = Join-Path ([IO.Path]::GetTempPath()) ('officeimo-roundtrip-' + [guid]::NewGuid().ToString('N'))
}

try {
    $results = Join-Path $OutputDirectory 'results'
    $failed = $false
    foreach ($framework in 'net472', 'net10.0') {
        $bin = Join-Path $OutputDirectory $framework
        dotnet build $project -c Release -f $framework -o $bin -nologo -v quiet | Out-Host
        if ($LASTEXITCODE -ne 0) { throw "Building the probe for $framework failed." }
        $exe = Join-Path $bin 'OfficeIMO.RoundTripNumberTextProbe'
        if ($framework -eq 'net472') { & "$exe.exe" $Count $results } else { dotnet "$exe.dll" $Count $results }
        if ($LASTEXITCODE -eq 2) { throw "The probe failed to start on $framework." }
        if ($LASTEXITCODE -ne 0) { $failed = $true }
    }
    if ($failed) { exit 1 }
} finally {
    if ($removeOutput -and (Test-Path $OutputDirectory)) { Remove-Item -LiteralPath $OutputDirectory -Recurse -Force }
}
