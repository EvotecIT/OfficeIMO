<#
.SYNOPSIS
Checks that .NET Framework writes the same round-trip number text as .NET Core.

.DESCRIPTION
Builds Build/RoundTripNumberText for net472 and net10.0, writes the text of the same seeded doubles and floats on
both, and fails when any line differs. .NET 10 uses the runtime's "R" format; .NET Framework uses the digits
OfficeRoundTripNumberText computes itself. Requires Windows for the net472 build. The check is opt-in: it takes a
minute or two for the default two million values.

.PARAMETER Count
How many doubles and floats to write.

.PARAMETER OutputDirectory
Where the build and the text files go; a new folder under the temporary directory by default, removed afterwards.
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
    foreach ($framework in 'net472', 'net10.0') {
        $bin = Join-Path $OutputDirectory $framework
        dotnet build $project -c Release -f $framework -o $bin -nologo -v quiet | Out-Host
        if ($LASTEXITCODE -ne 0) { throw "Building the probe for $framework failed." }
        $exe = Join-Path $bin 'OfficeIMO.RoundTripNumberTextProbe'
        if ($framework -eq 'net472') { & "$exe.exe" $Count $results } else { dotnet "$exe.dll" $Count $results }
        if ($LASTEXITCODE -ne 0) { throw "The probe failed on $framework." }
    }

    $failed = $false
    foreach ($kind in 'double', 'float') {
        $framework = [IO.File]::ReadAllLines((Join-Path $results "framework-$kind.txt"))
        $core = [IO.File]::ReadAllLines((Join-Path $results "core-$kind.txt"))
        $differences = 0
        for ($index = 0; $index -lt [Math]::Max($framework.Length, $core.Length); $index++) {
            if ($framework[$index] -cne $core[$index]) {
                if ($differences -lt 10) { Write-Host "$kind differs: framework '$($framework[$index])', core '$($core[$index])'" }
                $differences++
            }
        }
        if ($differences -gt 0) {
            Write-Host "$kind`: $differences of $($core.Length) lines differ."
            $failed = $true
        } else {
            Write-Host "$kind`: all $($core.Length) lines match."
        }
    }
    if ($failed) { exit 1 }
} finally {
    if ($removeOutput -and (Test-Path $OutputDirectory)) { Remove-Item -LiteralPath $OutputDirectory -Recurse -Force }
}
