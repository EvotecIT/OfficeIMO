<#
.SYNOPSIS
Checks OfficeIMO's number text and number parsing on .NET Framework against .NET 10.

.DESCRIPTION
Builds Build/RoundTripNumberText for net472 and net10.0 and runs three steps. The net10.0 run writes a parse corpus:
seeded digit strings of up to 25 digits across the exponent range, decimals at and next to the exact midpoint between
two doubles or two floats, formatted values and spellings, with .NET 10's double and float results. The net472 run
writes the text of seeded doubles and floats, fails when OfficeIMO's parser does not read every text back exactly,
fails when OfficeIMO's parser disagrees with .NET 10 on any corpus string, and times both parsers on the corpus. The
last net10.0 run fails unless every text written on .NET Framework is byte-identical to .NET 10's "R". Requires
Windows for the net472 build. The check is opt-in: it takes a few minutes for the default two million values.

.PARAMETER Count
How many doubles and floats to write, and how many random corpus strings to parse.

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
        dotnet build $project -c Release -f $framework -o (Join-Path $OutputDirectory $framework) -nologo -v quiet | Out-Host
        if ($LASTEXITCODE -ne 0) { throw "Building the probe for $framework failed." }
    }
    foreach ($step in @(@('net10.0', 'corpus'), @('net472', 'write'), @('net10.0', 'check'))) {
        $exe = Join-Path (Join-Path $OutputDirectory $step[0]) 'OfficeIMO.RoundTripNumberTextProbe'
        if ($step[0] -eq 'net472') { & "$exe.exe" $step[1] $Count $results } else { dotnet "$exe.dll" $step[1] $Count $results }
        if ($LASTEXITCODE -eq 2) { throw "The probe step $($step[1]) failed to start on $($step[0])." }
        if ($LASTEXITCODE -ne 0) { $failed = $true }
    }
    if ($failed) { exit 1 }
} finally {
    if ($removeOutput -and (Test-Path $OutputDirectory)) { Remove-Item -LiteralPath $OutputDirectory -Recurse -Force }
}
