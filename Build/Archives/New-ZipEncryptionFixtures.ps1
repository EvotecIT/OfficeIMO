#Requires -Version 7.0
<#
.SYNOPSIS
Regenerates the encrypted ZIP fixtures with Info-ZIP zip and 7-Zip.

.DESCRIPTION
Writes deterministic inputs, archives them with traditional ZipCrypto (Info-ZIP zip and 7-Zip) and WinZip AES
(7-Zip, 128, 192 and 256-bit keys), checks every archive with the producing tool's test mode, and records the tool
versions, commands, passwords and SHA-256 hashes in manifest.json. WinZip AES salts are random, so regenerated
archives differ byte for byte while decoding to the same inputs.
#>
[CmdletBinding()]
param(
    [string] $Destination = (Join-Path $PSScriptRoot '../../OfficeIMO.Archives.Tests/Fixtures/IndependentZipEncryption'),
    [string] $Zip = 'zip',
    [string] $SevenZip = '7zz'
)

$ErrorActionPreference = 'Stop'
$Destination = [System.IO.Path]::GetFullPath($Destination)
$inputs = Join-Path $Destination 'inputs'
New-Item -ItemType Directory -Force -Path $inputs | Out-Null

$text = [System.Text.StringBuilder]::new()
for ($line = 0; $line -lt 2400; $line++) {
    [void]$text.Append("Line $line of the harbour log: tide $($line % 17), wind $($line % 5) knots, buoy $($line % 41).`n")
}
[System.IO.File]::WriteAllBytes((Join-Path $inputs 'text.txt'), [System.Text.Encoding]::UTF8.GetBytes($text.ToString()))

$mixed = [System.Collections.Generic.List[byte]]::new()
[uint32] $state = 12345
for ($chunk = 0; $chunk -lt 48; $chunk++) {
    $mixed.AddRange([System.Text.Encoding]::UTF8.GetBytes(("Chunk $chunk repeats its own words. " * 40)))
    for ($index = 0; $index -lt 600; $index++) {
        $state = [uint32](([uint64]$state * 1103515245 + 12345) % 4294967296)
        $mixed.Add([byte](($state -shr 16) -band 0xff))
    }
}
[System.IO.File]::WriteAllBytes((Join-Path $inputs 'mixed.bin'), $mixed.ToArray())
[System.IO.File]::WriteAllBytes((Join-Path $inputs 'tiny.txt'), [System.Text.Encoding]::ASCII.GetBytes('short'))
[System.IO.File]::WriteAllBytes((Join-Path $inputs 'empty.bin'), [byte[]]::new(0))
$stamp = [datetime]::new(2026, 1, 2, 3, 4, 6)
Get-ChildItem $inputs -File | ForEach-Object { $_.LastWriteTime = $stamp }

$password = 'correct horse'
$members = @('text.txt', 'mixed.bin', 'tiny.txt', 'empty.bin')

function Invoke-Producer([string] $Tool, [string[]] $Arguments, [string] $Output) {
    $path = Join-Path $Destination $Output
    Remove-Item -LiteralPath $path -ErrorAction SilentlyContinue
    Push-Location $inputs
    try {
        & $Tool @Arguments $path @members | Out-Null
        if ($LASTEXITCODE -ne 0) { throw "$Tool failed for $Output." }
    } finally {
        Pop-Location
    }
}

# Info-ZIP reading standard input writes one member with a data descriptor, whose check byte is the time's high byte.
function Invoke-ZipFromStdin([string] $Output) {
    $path = Join-Path $Destination $Output
    Remove-Item -LiteralPath $path -ErrorAction SilentlyContinue
    $process = [System.Diagnostics.Process]::new()
    $process.StartInfo.FileName = (Get-Command $Zip).Source
    foreach ($argument in @('-q', '-P', $password, $path, '-')) { $process.StartInfo.ArgumentList.Add($argument) }
    $process.StartInfo.RedirectStandardInput = $true
    $process.StartInfo.UseShellExecute = $false
    [void]$process.Start()
    $bytes = [System.IO.File]::ReadAllBytes((Join-Path $inputs 'text.txt'))
    $process.StandardInput.BaseStream.Write($bytes, 0, $bytes.Length)
    $process.StandardInput.Close()
    $process.WaitForExit()
    if ($process.ExitCode -ne 0) { throw "$Zip failed for $Output." }
}

$fixtures = @(
    @{ File = 'zipcrypto-infozip.zip'; Tool = $Zip; Arguments = @('-q', '-X', '-P', $password); Password = $password }
    @{ File = 'zipcrypto-infozip-stored.zip'; Tool = $Zip; Arguments = @('-q', '-X', '-0', '-P', $password); Password = $password }
    @{ File = 'zipcrypto-7zip.zip'; Tool = $SevenZip; Arguments = @('a', '-tzip', "-p$password", '-mem=ZipCrypto', '-mx=5'); Password = $password }
    @{ File = 'aes128-7zip.zip'; Tool = $SevenZip; Arguments = @('a', '-tzip', "-p$password", '-mem=AES128', '-mx=5'); Password = $password }
    @{ File = 'aes192-7zip.zip'; Tool = $SevenZip; Arguments = @('a', '-tzip', "-p$password", '-mem=AES192', '-mx=9'); Password = $password }
    @{ File = 'aes256-7zip.zip'; Tool = $SevenZip; Arguments = @('a', '-tzip', "-p$password", '-mem=AES256', '-mx=5'); Password = $password }
    @{ File = 'aes256-7zip-stored.zip'; Tool = $SevenZip; Arguments = @('a', '-tzip', "-p$password", '-mem=AES256', '-mx=0'); Password = $password }
)

$records = foreach ($fixture in $fixtures) {
    Invoke-Producer $fixture.Tool $fixture.Arguments $fixture.File
    $shown = $fixture.Arguments | ForEach-Object { $_ -replace [regex]::Escape($fixture.Password), '<password>' }
    [ordered]@{
        file = $fixture.File
        password = $fixture.Password
        command = "$([System.IO.Path]::GetFileName($fixture.Tool)) $($shown -join ' ') $($fixture.File) $($members -join ' ')"
    }
}
Invoke-ZipFromStdin 'zipcrypto-infozip-descriptor.zip'
$records += [ordered]@{ file = 'zipcrypto-infozip-descriptor.zip'; password = $password; command = 'zip -q -P <password> zipcrypto-infozip-descriptor.zip - < inputs/text.txt' }

foreach ($record in $records) {
    $path = Join-Path $Destination $record.file
    & $SevenZip t "-p$($record.password)" $path | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "7-Zip rejected $($record.file)." }
    $record.sha256 = (Get-FileHash -Algorithm SHA256 $path).Hash.ToLowerInvariant()
}

$manifest = [ordered]@{
    producers = @(((& $Zip -v) | Select-Object -Skip 1 -First 1).Trim(), ((& $SevenZip) | Select-Object -Skip 1 -First 1).Trim())
    sources = @('https://infozip.sourceforge.net/', 'https://www.7-zip.org/')
    verification = '7-Zip test mode (7zz t) with the recorded password applied to each fixture.'
    fixtures = $records
}
$manifest | ConvertTo-Json -Depth 4 | Set-Content -Path (Join-Path $Destination 'manifest.json') -Encoding utf8NoBOM
Write-Host "Wrote $($records.Count) fixtures to $Destination"
