#Requires -Version 7.0
<#
.SYNOPSIS
Regenerates the compressed TAR fixtures with bsdtar, zstd, xz and lzip.

.DESCRIPTION
Writes a small deterministic tree (a text file, a binary file and a symbolic link), archives it with bsdtar as pax,
pipes the archive through each reference compressor, checks every output with the producing tool's test mode and
records the tool versions, commands and SHA-256 hashes in manifest.json. The uncompressed archive.tar is kept as the
reference content; the input tree is removed.
#>
[CmdletBinding()]
param(
    [string] $Destination = (Join-Path $PSScriptRoot '../../OfficeIMO.Archives.Tests/Fixtures/IndependentTarWrappers'),
    [string] $Tar = 'bsdtar',
    [string] $Zstd = 'zstd',
    [string] $Xz = 'xz',
    [string] $Lzip = 'lzip'
)

$ErrorActionPreference = 'Stop'
$Destination = [System.IO.Path]::GetFullPath($Destination)
$tree = Join-Path $Destination 'inputs'
$docs = Join-Path $tree 'docs'
Remove-Item -Recurse -Force -LiteralPath $tree -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force -Path $docs | Out-Null

$text = [System.Text.StringBuilder]::new()
for ($line = 0; $line -lt 2400; $line++) {
    [void]$text.Append("Line $line of the harbour log: tide $($line % 17), wind $($line % 5) knots, buoy $($line % 41).`n")
}
[System.IO.File]::WriteAllBytes((Join-Path $docs 'text.txt'), [System.Text.Encoding]::UTF8.GetBytes($text.ToString()))
$mixed = [System.Collections.Generic.List[byte]]::new()
[uint32] $state = 12345
for ($chunk = 0; $chunk -lt 48; $chunk++) {
    $mixed.AddRange([System.Text.Encoding]::UTF8.GetBytes(("Chunk $chunk repeats its own words. " * 40)))
    for ($index = 0; $index -lt 600; $index++) {
        $state = [uint32](([uint64]$state * 1103515245 + 12345) % 4294967296)
        $mixed.Add([byte](($state -shr 16) -band 0xff))
    }
}
[System.IO.File]::WriteAllBytes((Join-Path $docs 'mixed.bin'), $mixed.ToArray())
New-Item -ItemType SymbolicLink -Path (Join-Path $docs 'text-link.txt') -Target 'text.txt' | Out-Null
$stamp = [datetime]::new(2026, 1, 2, 3, 4, 6)
Get-ChildItem $docs -File | ForEach-Object { $_.LastWriteTime = $stamp }

$archive = Join-Path $Destination 'archive.tar'
$env:COPYFILE_DISABLE = '1'
Push-Location $tree
try {
    & $Tar --format pax --no-xattrs --no-mac-metadata --uid 0 --gid 0 --uname '' --gname '' -cf $archive docs/text.txt docs/mixed.bin docs/text-link.txt
    if ($LASTEXITCODE -ne 0) { throw 'bsdtar failed.' }
} finally {
    Pop-Location
}

# The archive is piped, so each compressor sees no input size and keeps its configured window.
function Invoke-Compressor([string] $Tool, [string[]] $Arguments, [string] $Output) {
    $process = [System.Diagnostics.Process]::new()
    $process.StartInfo.FileName = (Get-Command $Tool).Source
    foreach ($argument in $Arguments + @('-c')) { $process.StartInfo.ArgumentList.Add($argument) }
    $process.StartInfo.RedirectStandardInput = $true
    $process.StartInfo.RedirectStandardOutput = $true
    $process.StartInfo.UseShellExecute = $false
    [void]$process.Start()
    $file = [System.IO.File]::Create((Join-Path $Destination $Output))
    $copy = $process.StandardOutput.BaseStream.CopyToAsync($file)
    $bytes = [System.IO.File]::ReadAllBytes($archive)
    $process.StandardInput.BaseStream.Write($bytes, 0, $bytes.Length)
    $process.StandardInput.Close()
    $copy.Wait()
    $file.Dispose()
    $process.WaitForExit()
    if ($process.ExitCode -ne 0) { throw "$Tool failed for $Output." }
}

$fixtures = @(
    @{ File = 'archive.tar.zst'; Tool = $Zstd; Arguments = @('-19', '-q') }
    @{ File = 'archive-window128m.tar.zst'; Tool = $Zstd; Arguments = @('-3', '--long=27', '-q') }
    @{ File = 'archive.tar.xz'; Tool = $Xz; Arguments = @('-6') }
    @{ File = 'archive.tar.lz'; Tool = $Lzip; Arguments = @('-6') }
)
$records = @([ordered]@{
    file = 'archive.tar'
    command = 'bsdtar --format pax --no-xattrs --no-mac-metadata --uid 0 --gid 0 --uname "" --gname "" -cf archive.tar docs/text.txt docs/mixed.bin docs/text-link.txt'
})
foreach ($fixture in $fixtures) {
    Invoke-Compressor $fixture.Tool $fixture.Arguments $fixture.File
    $records += [ordered]@{
        file = $fixture.File
        command = "$([System.IO.Path]::GetFileName($fixture.Tool)) $($fixture.Arguments -join ' ') -c < archive.tar"
    }
}

foreach ($record in $records) {
    $path = Join-Path $Destination $record.file
    switch -Wildcard ($record.file) {
        '*.zst' { & $Zstd -t -q --long=27 $path }
        '*.xz' { & $Xz -t $path }
        '*.lz' { & $Lzip -t $path }
        default { & $Tar -tf $path | Out-Null }
    }
    if ($LASTEXITCODE -ne 0) { throw "The producing tool rejected $($record.file)." }
    $record.sha256 = (Get-FileHash -Algorithm SHA256 $path).Hash.ToLowerInvariant()
}

$manifest = [ordered]@{
    producers = @(
        ((& $Tar --version) | Select-Object -First 1),
        ((& $Zstd --version) | Select-Object -First 1),
        ((& $Xz --version) | Select-Object -First 1),
        ((& $Lzip --version) | Select-Object -First 1))
    verification = 'bsdtar -t, zstd -t, xz -t or lzip -t applied to each fixture.'
    fixtures = $records
}
# The tree holds a symbolic link, which is awkward to keep in Git on Windows; archive.tar carries the same content.
Remove-Item -Recurse -Force -LiteralPath $tree
$manifest | ConvertTo-Json -Depth 4 | Set-Content -Path (Join-Path $Destination 'manifest.json') -Encoding utf8NoBOM
Write-Host "Wrote $($records.Count) fixtures to $Destination"
