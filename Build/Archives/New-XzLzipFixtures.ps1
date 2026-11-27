#Requires -Version 7.0
<#
.SYNOPSIS
Regenerates the independent xz and lzip fixtures with the reference xz and lzip tools.

.DESCRIPTION
Writes deterministic inputs, compresses them with xz (XZ Utils) and lzip using the options each fixture names,
verifies every output with the producing tool's test mode, and records the tool versions, commands and SHA-256
hashes in manifest.json.
#>
[CmdletBinding()]
param(
    [string] $Destination = (Join-Path $PSScriptRoot '../../OfficeIMO.Archives.Tests/Fixtures/IndependentXzLzip'),
    [string] $Xz = 'xz',
    [string] $Lzip = 'lzip'
)

$ErrorActionPreference = 'Stop'
$Destination = [System.IO.Path]::GetFullPath($Destination)
New-Item -ItemType Directory -Force -Path $Destination | Out-Null
$inputs = Join-Path $Destination 'inputs'
New-Item -ItemType Directory -Force -Path $inputs | Out-Null

function Write-Bytes([string] $Name, [byte[]] $Bytes) {
    [System.IO.File]::WriteAllBytes((Join-Path $inputs $Name), $Bytes)
}

$text = [System.Text.StringBuilder]::new()
for ($line = 0; $line -lt 2400; $line++) {
    [void]$text.Append("Line $line of the harbour log: tide $($line % 17), wind $($line % 5) knots, buoy $($line % 41).`n")
}
Write-Bytes 'text.txt' ([System.Text.Encoding]::UTF8.GetBytes($text.ToString()))

$mixed = [System.Collections.Generic.List[byte]]::new()
[uint32] $state = 12345
for ($chunk = 0; $chunk -lt 48; $chunk++) {
    $mixed.AddRange([System.Text.Encoding]::UTF8.GetBytes(("Chunk $chunk repeats its own words. " * 40)))
    for ($index = 0; $index -lt 600; $index++) {
        $state = [uint32](([uint64]$state * 1103515245 + 12345) % 4294967296)
        $mixed.Add([byte](($state -shr 16) -band 0xff))
    }
}
Write-Bytes 'mixed.bin' $mixed.ToArray()

# 16-bit samples of a slow ramp, the kind of data the delta filter is for.
$samples = [byte[]]::new(65536)
for ($index = 0; $index -lt 32768; $index++) {
    $value = [int](($index * 3) % 65536)
    $samples[2 * $index] = [byte]($value -band 0xff)
    $samples[2 * $index + 1] = [byte]($value -shr 8)
}
Write-Bytes 'samples.bin' $samples
Write-Bytes 'empty.bin' ([byte[]]::new(0))

function Invoke-Tool([string] $Tool, [string[]] $Arguments, [string] $InputName, [string] $Output) {
    Push-Location $Destination
    try {
        $bytes = [System.IO.File]::ReadAllBytes((Join-Path $inputs $InputName))
        $process = [System.Diagnostics.Process]::new()
        $process.StartInfo.FileName = (Get-Command $Tool).Source
        foreach ($argument in $Arguments + @('-c')) { $process.StartInfo.ArgumentList.Add($argument) }
        $process.StartInfo.RedirectStandardInput = $true
        $process.StartInfo.RedirectStandardOutput = $true
        $process.StartInfo.UseShellExecute = $false
        [void]$process.Start()
        $copy = $process.StandardOutput.BaseStream.CopyToAsync(($file = [System.IO.File]::Create((Join-Path $Destination $Output))))
        $process.StandardInput.BaseStream.Write($bytes, 0, $bytes.Length)
        $process.StandardInput.Close()
        $copy.Wait()
        $file.Dispose()
        $process.WaitForExit()
        if ($process.ExitCode -ne 0) { throw "$Tool failed for $Output." }
    } finally {
        Pop-Location
    }
}

# Input is piped, so the tools see no file size and keep their configured dictionaries and block layouts.
$fixtures = @(
    @{ File = 'text-preset0.xz'; Tool = $Xz; Input = 'text.txt'; Arguments = @('-0') }
    @{ File = 'text-preset6-crc64.xz'; Tool = $Xz; Input = 'text.txt'; Arguments = @('-6') }
    @{ File = 'text-crc32.xz'; Tool = $Xz; Input = 'text.txt'; Arguments = @('-6', '--check=crc32') }
    @{ File = 'text-sha256.xz'; Tool = $Xz; Input = 'text.txt'; Arguments = @('-6', '--check=sha256') }
    @{ File = 'text-nocheck.xz'; Tool = $Xz; Input = 'text.txt'; Arguments = @('-6', '--check=none') }
    @{ File = 'mixed-preset9e.xz'; Tool = $Xz; Input = 'mixed.bin'; Arguments = @('-9e') }
    @{ File = 'mixed-blocks.xz'; Tool = $Xz; Input = 'mixed.bin'; Arguments = @('-6', '--block-size=16KiB') }
    @{ File = 'mixed-threads.xz'; Tool = $Xz; Input = 'mixed.bin'; Arguments = @('-6', '-T2', '--block-size=32KiB') }
    @{ File = 'samples-delta.xz'; Tool = $Xz; Input = 'samples.bin'; Arguments = @('--delta=dist=2', '--lzma2=preset=6') }
    @{ File = 'text-x86.xz'; Tool = $Xz; Input = 'text.txt'; Arguments = @('--x86', '--lzma2=preset=6') }
    @{ File = 'text-dict128m.xz'; Tool = $Xz; Input = 'text.txt'; Arguments = @('--lzma2=preset=6,dict=128MiB') }
    @{ File = 'text-dict4k.xz'; Tool = $Xz; Input = 'text.txt'; Arguments = @('--lzma2=preset=6,dict=4KiB') }
    @{ File = 'empty.xz'; Tool = $Xz; Input = 'empty.bin'; Arguments = @('-6') }
    @{ File = 'text-level0.lz'; Tool = $Lzip; Input = 'text.txt'; Arguments = @('-0') }
    @{ File = 'text-level6.lz'; Tool = $Lzip; Input = 'text.txt'; Arguments = @('-6') }
    @{ File = 'mixed-level9.lz'; Tool = $Lzip; Input = 'mixed.bin'; Arguments = @('-9') }
    @{ File = 'text-dict4k.lz'; Tool = $Lzip; Input = 'text.txt'; Arguments = @('-6', '--dictionary-size=4KiB') }
    @{ File = 'empty.lz'; Tool = $Lzip; Input = 'empty.bin'; Arguments = @('-6') }
)

$records = foreach ($fixture in $fixtures) {
    Invoke-Tool $fixture.Tool $fixture.Arguments $fixture.Input $fixture.File
    $name = [System.IO.Path]::GetFileNameWithoutExtension($fixture.Tool)
    [ordered]@{
        file = $fixture.File
        input = "inputs/$($fixture.Input)"
        command = "$name $($fixture.Arguments -join ' ') -c < inputs/$($fixture.Input)"
    }
}

# Two xz streams with four bytes of stream padding between them.
$first = [System.IO.File]::ReadAllBytes((Join-Path $Destination 'text-preset0.xz'))
$second = [System.IO.File]::ReadAllBytes((Join-Path $Destination 'mixed-blocks.xz'))
[System.IO.File]::WriteAllBytes((Join-Path $Destination 'concatenated.xz'), [byte[]]($first + [byte[]]::new(4) + $second))
$records += [ordered]@{ file = 'concatenated.xz'; input = 'inputs/text.txt + inputs/mixed.bin'; command = 'text-preset0.xz, four null bytes of stream padding, then mixed-blocks.xz' }

$first = [System.IO.File]::ReadAllBytes((Join-Path $Destination 'text-level0.lz'))
$second = [System.IO.File]::ReadAllBytes((Join-Path $Destination 'mixed-level9.lz'))
[System.IO.File]::WriteAllBytes((Join-Path $Destination 'concatenated.lz'), [byte[]]($first + $second))
$records += [ordered]@{ file = 'concatenated.lz'; input = 'inputs/text.txt + inputs/mixed.bin'; command = 'text-level0.lz followed by mixed-level9.lz' }

foreach ($record in $records) {
    $path = Join-Path $Destination $record.file
    if ($record.file.EndsWith('.xz')) { & $Xz -t --memlimit=0 $path } else { & $Lzip -t $path }
    if ($LASTEXITCODE -ne 0) { throw "The producing tool rejected $($record.file)." }
    $record.sha256 = (Get-FileHash -Algorithm SHA256 $path).Hash.ToLowerInvariant()
}

$manifest = [ordered]@{
    producers = @(((& $Xz --version) | Select-Object -First 1), ((& $Lzip --version) | Select-Object -First 1))
    sources = @('https://tukaani.org/xz/', 'https://www.nongnu.org/lzip/')
    verification = 'xz -t or lzip -t applied to each fixture.'
    fixtures = $records
}
$manifest | ConvertTo-Json -Depth 4 | Set-Content -Path (Join-Path $Destination 'manifest.json') -Encoding utf8NoBOM
Write-Host "Wrote $($records.Count) fixtures to $Destination"
