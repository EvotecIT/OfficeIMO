#Requires -Version 7.4
<#
.SYNOPSIS
Regenerates the independent Zstandard fixtures with the reference zstd command-line tool.

.DESCRIPTION
Writes deterministic inputs, compresses them with zstd using the options each fixture names, verifies every output
with `zstd -t`, and records the tool version, commands and SHA-256 hashes in manifest.json. The fixtures exercise
frames written by an independent producer: levels, block splitting, small windows that force history sliding,
large declared windows, multiple frames, skippable frames, checksums and dictionaries.
#>
[CmdletBinding()]
param(
    [string] $Destination = (Join-Path $PSScriptRoot '../../OfficeIMO.Archives.Tests/Fixtures/IndependentZstandard'),
    [string] $Zstd = 'zstd'
)

$ErrorActionPreference = 'Stop'
# Resolve the tool once, so a relative path keeps working after the fixture loop changes directory.
$Zstd = (Get-Command -Name $Zstd -CommandType Application -ErrorAction Stop | Select-Object -First 1).Source
$Destination = [System.IO.Path]::GetFullPath($Destination)
New-Item -ItemType Directory -Force -Path $Destination | Out-Null
$inputs = Join-Path $Destination 'inputs'
New-Item -ItemType Directory -Force -Path $inputs | Out-Null

function Write-Bytes([string] $Name, [byte[]] $Bytes) {
    [System.IO.File]::WriteAllBytes((Join-Path $inputs $Name), $Bytes)
}

# Text with long and short repeats.
$text = [System.Text.StringBuilder]::new()
for ($line = 0; $line -lt 2400; $line++) {
    [void]$text.Append("Line $line of the harbour log: tide $($line % 17), wind $($line % 5) knots, buoy $($line % 41).`n")
}
Write-Bytes 'text.txt' ([System.Text.Encoding]::UTF8.GetBytes($text.ToString()))

# Text interleaved with pseudo-random bytes from a fixed linear congruential generator.
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

Write-Bytes 'empty.bin' ([byte[]]::new(0))

# Small JSON-like messages for dictionary training and a separate message compressed against it.
$samples = Join-Path ([System.IO.Path]::GetTempPath()) ("zstd-samples-" + [guid]::NewGuid().ToString('n'))
New-Item -ItemType Directory -Path $samples | Out-Null
try {
    for ($index = 0; $index -lt 400; $index++) {
        $message = "{""id"":$index,""kind"":""move"",""x"":$($index * 7 % 100),""y"":$($index * 13 % 100),""owner"":""player-$($index % 4)""}"
        [System.IO.File]::WriteAllText((Join-Path $samples "$index.json"), $message)
    }
    $message = '{"id":9001,"kind":"move","x":42,"y":17,"owner":"player-1","note":"the message compressed against the dictionary"}'
    Write-Bytes 'message.json' ([System.Text.Encoding]::UTF8.GetBytes($message))
    & $Zstd -q -f --train (Get-ChildItem $samples -File).FullName --maxdict=1024 --dictID=4242 -o (Join-Path $Destination 'trained.dict')
    if ($LASTEXITCODE -ne 0) { throw 'zstd --train failed.' }
} finally {
    Remove-Item -Recurse -Force $samples
}
# A raw-content dictionary: any bytes without the dictionary magic number.
[System.IO.File]::WriteAllBytes((Join-Path $Destination 'raw-content.dict'),
    [System.Text.Encoding]::UTF8.GetBytes('{"id":0,"kind":"move","x":0,"y":0,"owner":"player-0"} the message compressed against the dictionary'))

$fixtures = @(
    @{ File = 'text-level1.zst'; Input = 'text.txt'; Arguments = @('-1') }
    @{ File = 'text-level3-check.zst'; Input = 'text.txt'; Arguments = @('-3', '--check') }
    @{ File = 'text-level19-nocheck.zst'; Input = 'text.txt'; Arguments = @('-19', '--no-check') }
    @{ File = 'mixed-level22.zst'; Input = 'mixed.bin'; Arguments = @('--ultra', '-22') }
    @{ File = 'mixed-level-fast.zst'; Input = 'mixed.bin'; Arguments = @('--fast=5') }
    @{ File = 'mixed-window1k.zst'; Input = 'mixed.bin'; Arguments = @('-12', '--zstd=windowLog=10') }
    @{ File = 'mixed-blocks.zst'; Input = 'mixed.bin'; Arguments = @('-6', '--zstd=targetLength=16', '--no-content-size') }
    @{ File = 'empty.zst'; Input = 'empty.bin'; Arguments = @('-3') }
    @{ File = 'message-trained-dict.zst'; Input = 'message.json'; Arguments = @('-3', '-D', 'trained.dict') }
    @{ File = 'message-raw-dict.zst'; Input = 'message.json'; Arguments = @('-3', '-D', 'raw-content.dict') }
)

$records = foreach ($fixture in $fixtures) {
    $output = Join-Path $Destination $fixture.File
    Push-Location $Destination
    try {
        & $Zstd -q -f @($fixture.Arguments) (Join-Path 'inputs' $fixture.Input) -o $fixture.File
        if ($LASTEXITCODE -ne 0) { throw "zstd failed for $($fixture.File)." }
    } finally {
        Pop-Location
    }
    [ordered]@{
        file = $fixture.File
        input = "inputs/$($fixture.Input)"
        command = "zstd $($fixture.Arguments -join ' ') inputs/$($fixture.Input) -o $($fixture.File)"
    }
}

# A frame whose size is unknown keeps its declared 128 MiB window, which is above the default limit.
Push-Location $Destination
try {
    Get-Content -AsByteStream -Raw (Join-Path 'inputs' 'text.txt') | & $Zstd -q -3 --long=27 -c > 'text-window128m.zst'
} finally {
    Pop-Location
}
$records += [ordered]@{ file = 'text-window128m.zst'; input = 'inputs/text.txt'; command = 'zstd -3 --long=27 -c < inputs/text.txt' }

# Four copies of the text through a 1 KiB window. The output passes the decoder's first history buffer twice over, so
# matches after each slide read history that was moved.
$repeated = Join-Path ([System.IO.Path]::GetTempPath()) ("zstd-repeated-" + [guid]::NewGuid().ToString('n'))
try {
    $once = [System.IO.File]::ReadAllBytes((Join-Path $inputs 'text.txt'))
    [System.IO.File]::WriteAllBytes($repeated, [byte[]]($once + $once + $once + $once))
    & $Zstd -q -f -12 --zstd=windowLog=10 $repeated -o (Join-Path $Destination 'text4-window1k.zst')
    if ($LASTEXITCODE -ne 0) { throw 'zstd failed for text4-window1k.zst.' }
} finally {
    Remove-Item -Force -ErrorAction SilentlyContinue $repeated
}
$records += [ordered]@{
    file = 'text4-window1k.zst'
    input = 'inputs/text.txt four times'
    command = 'zstd -12 --zstd=windowLog=10 on inputs/text.txt repeated four times'
}

# Two frames from separate runs, with a skippable frame between them.
$skippable = [byte[]](0x50, 0x2A, 0x4D, 0x18, 0x05, 0x00, 0x00, 0x00, 0x68, 0x65, 0x6C, 0x6C, 0x6F)
$first = [System.IO.File]::ReadAllBytes((Join-Path $Destination 'text-level1.zst'))
$second = [System.IO.File]::ReadAllBytes((Join-Path $Destination 'mixed-level-fast.zst'))
[System.IO.File]::WriteAllBytes((Join-Path $Destination 'concatenated.zst'), [byte[]]($first + $skippable + $second))
$records += [ordered]@{
    file = 'concatenated.zst'
    input = 'inputs/text.txt + inputs/mixed.bin'
    command = 'text-level1.zst, a 5-byte skippable frame (magic 0x184D2A50), then mixed-level-fast.zst'
}

$version = (& $Zstd --version) -join ' '
foreach ($record in $records) {
    $path = Join-Path $Destination $record.file
    $dictionary = if ($record.file -like '*trained-dict*') { @('-D', (Join-Path $Destination 'trained.dict')) }
        elseif ($record.file -like '*raw-dict*') { @('-D', (Join-Path $Destination 'raw-content.dict')) }
        else { @() }
    & $Zstd -q -t --long=27 @dictionary $path
    if ($LASTEXITCODE -ne 0) { throw "zstd -t rejected $($record.file)." }
    $record.sha256 = (Get-FileHash -Algorithm SHA256 $path).Hash.ToLowerInvariant()
}

$manifest = [ordered]@{
    producer = $version
    source = 'https://github.com/facebook/zstd/releases'
    verification = 'zstd -t (with -D for dictionary fixtures and --long=27 for the large window) applied to each fixture.'
    dictionaries = [ordered]@{
        'trained.dict' = 'zstd --train over 400 generated JSON messages, --maxdict=1024 --dictID=4242.'
        'raw-content.dict' = 'Plain bytes without the dictionary magic number.'
    }
    fixtures = $records
}
$manifest | ConvertTo-Json -Depth 4 | Set-Content -Path (Join-Path $Destination 'manifest.json') -Encoding utf8NoBOM
Write-Host "Wrote $($records.Count) fixtures to $Destination"
