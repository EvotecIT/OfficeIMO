[CmdletBinding()]
param([Parameter(Mandatory)][string] $OutputPath)

$ErrorActionPreference = 'Stop'
if (-not $IsWindows) { throw 'The native Windows Shell producer requires Windows.' }
$OutputPath = [IO.Path]::GetFullPath($OutputPath)
if (Test-Path -LiteralPath $OutputPath) { throw 'Use a fresh task-owned output path.' }
$null = New-Item -ItemType Directory -Path $OutputPath
$source = Join-Path $OutputPath 'source'
$null = New-Item -ItemType Directory -Path $source
$names = @('Zażółć-gęślą-jaźń.txt', 'Łódź.txt')
$payload = [Text.Encoding]::UTF8.GetBytes('Zażółć gęślą jaźń — native Windows Shell ZIP')
foreach ($name in $names) { [IO.File]::WriteAllBytes((Join-Path $source $name), $payload) }
$zipPath = Join-Path $OutputPath 'windows-shell-polish.zip'
[IO.File]::WriteAllBytes($zipPath, [byte[]](0x50, 0x4b, 5, 6) + [byte[]]::new(18))
$shell = New-Object -ComObject Shell.Application
try {
    $destination = $shell.NameSpace($zipPath)
    $folder = $shell.NameSpace($source)
    if ($null -eq $destination -or $null -eq $folder) { throw 'Windows compressed-folder namespace unavailable.' }
    # Use Explorer's compressed-folder handler, not a managed ZIP writer or Compress-Archive.
    $items = $folder.Items()
    $destination.CopyHere($items, 0x614)
    $deadline = [DateTime]::UtcNow.AddSeconds(60)
    do {
        Start-Sleep -Milliseconds 250
        try {
            $read = [IO.Compression.ZipFile]::OpenRead($zipPath)
            try { $complete = $read.Entries.Count -eq $names.Count -and @($read.Entries | Where-Object Length -ne $payload.Length).Count -eq 0 }
            finally { $read.Dispose() }
        } catch [IO.IOException], [IO.InvalidDataException] { $complete = $false }
    } until ($complete -or [DateTime]::UtcNow -ge $deadline)
    if (-not $complete) { throw 'Native Shell ZIP production did not complete.' }
} finally {
    if ($null -ne $items) { $null = [Runtime.InteropServices.Marshal]::FinalReleaseComObject($items) }
    if ($null -ne $destination) { $null = [Runtime.InteropServices.Marshal]::FinalReleaseComObject($destination) }
    if ($null -ne $folder) { $null = [Runtime.InteropServices.Marshal]::FinalReleaseComObject($folder) }
    $null = [Runtime.InteropServices.Marshal]::FinalReleaseComObject($shell)
}

# Record exact raw central-directory names and choose only an encoding that reproduces all authored names.
$bytes = [IO.File]::ReadAllBytes($zipPath)
$end = $bytes.Length - 22
while ($end -ge 0 -and [BitConverter]::ToUInt32($bytes, $end) -ne 0x06054b50) { $end-- }
if ($end -lt 0) { throw 'Missing ZIP end record.' }
$offset = [int][BitConverter]::ToUInt32($bytes, $end + 16)
$raw = @(); $flags = @()
for ($index = 0; $index -lt $names.Count; $index++) {
    if ([BitConverter]::ToUInt32($bytes, $offset) -ne 0x02014b50) { throw 'Malformed central directory.' }
    $length = [BitConverter]::ToUInt16($bytes, $offset + 28)
    $raw += ,$bytes[$($offset + 46)..$($offset + 45 + $length)]
    $flags += [BitConverter]::ToUInt16($bytes, $offset + 8)
    $offset += 46 + $length + [BitConverter]::ToUInt16($bytes, $offset + 30) + [BitConverter]::ToUInt16($bytes, $offset + 32)
}
[Text.Encoding]::RegisterProvider([Text.CodePagesEncodingProvider]::Instance)
$policy = $null
$decoded = @()
foreach ($page in @(65001, [Globalization.CultureInfo]::CurrentCulture.TextInfo.ANSICodePage, 1250) | Select-Object -Unique) {
    $encoding = [Text.Encoding]::GetEncoding($page, [Text.EncoderExceptionFallback]::new(), [Text.DecoderExceptionFallback]::new())
    try { $candidate = @($raw | ForEach-Object { $encoding.GetString([byte[]]$_) }) } catch [Text.DecoderFallbackException] { continue }
    if (@(Compare-Object ($names | Sort-Object) ($candidate | Sort-Object)).Count -eq 0) {
        $decoded = $candidate
        if (@($flags | Where-Object { ($_ -band 0x800) -eq 0 }).Count -gt 0) { $policy = $page }
        break
    }
}
if ($decoded.Count -ne $names.Count) { throw 'Native producer names do not match the authored Polish filenames.' }
$sha = [Security.Cryptography.SHA256]::Create()
try { $payloadHash = [BitConverter]::ToString($sha.ComputeHash($payload)).Replace('-', '').ToLowerInvariant() }
finally { $sha.Dispose() }
$producer = Join-Path $env:WINDIR 'System32/zipfldr.dll'
@{
    file = [IO.Path]::GetFileName($zipPath); sha256 = (Get-FileHash $zipPath -Algorithm SHA256).Hash.ToLowerInvariant()
    legacyCodePage = $policy
    entries = @($decoded | ForEach-Object { @{ name = $_; sha256 = $payloadHash; length = $payload.Length } })
    producer = @{ mechanism = 'Shell.Application compressed-folder CopyHere'; windows = [Environment]::OSVersion.VersionString
        zipfldrVersion = (Get-Item $producer).VersionInfo.FileVersion; zipfldrSha256 = (Get-FileHash $producer -Algorithm SHA256).Hash.ToLowerInvariant()
        explorerVersion = (Get-Item (Join-Path $env:WINDIR 'explorer.exe')).VersionInfo.FileVersion
        interactiveExplorerUI = $false; rawNames = @($raw | ForEach-Object { [Convert]::ToBase64String([byte[]]$_) }); flags = $flags }
} | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $OutputPath 'fixture.json') -Encoding utf8
