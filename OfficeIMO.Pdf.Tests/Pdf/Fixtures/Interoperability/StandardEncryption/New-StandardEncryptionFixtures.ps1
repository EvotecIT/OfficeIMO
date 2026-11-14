#Requires -Version 7.4
<#
.SYNOPSIS
Writes the Standard security fixtures in this folder with .NET cryptography only (no OfficeIMO code).

.DESCRIPTION
Each fixture is a one-page PDF whose /Info /Title and page content stream are encrypted.
The password algorithms follow ISO 32000-2 7.6.4.3.3 (algorithm 2.B) and 7.6.4.4.7-10; the
revision 7 file uses AES-256-GCM crypt filters (/AESV4, /V 6, /R 7) as ISO/TS 32003 describes,
with System.Security.Cryptography.AesGcm. Run it with PowerShell 7.4 or later from any folder.
#>
[CmdletBinding()]
param([string] $OutputDirectory = $PSScriptRoot)

$ErrorActionPreference = 'Stop'
$latin1 = [System.Text.Encoding]::Latin1

function Get-RandomBytes([int] $Count) { [System.Security.Cryptography.RandomNumberGenerator]::GetBytes($Count) }

function Join-Bytes([byte[][]] $Parts) {
    $stream = [System.IO.MemoryStream]::new()
    foreach ($part in $Parts) { if ($null -ne $part) { $stream.Write($part, 0, $part.Length) } }
    return , $stream.ToArray()
}

function Invoke-Aes([byte[]] $Key, [byte[]] $Iv, [byte[]] $Data, [System.Security.Cryptography.CipherMode] $Mode, [System.Security.Cryptography.PaddingMode] $Padding) {
    $aes = [System.Security.Cryptography.Aes]::Create()
    try {
        $aes.Key = $Key
        if ($Mode -eq 'ECB') { return , $aes.EncryptEcb($Data, $Padding) }
        return , $aes.EncryptCbc($Data, $Iv, $Padding)
    } finally { $aes.Dispose() }
}

# ISO 32000-2 algorithm 2.B (revision 6 and later password hash).
function Get-Revision6Hash([byte[]] $Password, [byte[]] $Salt, [byte[]] $UserData) {
    $k = [System.Security.Cryptography.SHA256]::HashData((Join-Bytes @($Password, $Salt, $UserData)))
    $round = 0
    while ($true) {
        $block = Join-Bytes @($Password, $k, $UserData)
        $k1 = [byte[]]::new($block.Length * 64)
        for ($i = 0; $i -lt 64; $i++) { [Array]::Copy($block, 0, $k1, $i * $block.Length, $block.Length) }
        $e = Invoke-Aes $k[0..15] $k[16..31] $k1 CBC None
        $sum = 0
        for ($i = 0; $i -lt 16; $i++) { $sum += $e[$i] }
        switch ($sum % 3) {
            0 { $k = [System.Security.Cryptography.SHA256]::HashData($e) }
            1 { $k = [System.Security.Cryptography.SHA384]::HashData($e) }
            2 { $k = [System.Security.Cryptography.SHA512]::HashData($e) }
        }
        $round++
        if ($round -ge 64 -and $e[$e.Length - 1] -le ($round - 32)) { break }
    }
    return , [byte[]]$k[0..31]
}

function New-EncryptionEntries([byte[]] $UserPassword, [byte[]] $OwnerPassword, [byte[]] $FileKey, [int] $Permissions) {
    $userValidationSalt = Get-RandomBytes 8; $userKeySalt = Get-RandomBytes 8
    $u = Join-Bytes @((Get-Revision6Hash $UserPassword $userValidationSalt @()), $userValidationSalt, $userKeySalt)
    $ue = Invoke-Aes (Get-Revision6Hash $UserPassword $userKeySalt @()) ([byte[]]::new(16)) $FileKey CBC None
    $ownerValidationSalt = Get-RandomBytes 8; $ownerKeySalt = Get-RandomBytes 8
    $o = Join-Bytes @((Get-Revision6Hash $OwnerPassword $ownerValidationSalt $u), $ownerValidationSalt, $ownerKeySalt)
    $oe = Invoke-Aes (Get-Revision6Hash $OwnerPassword $ownerKeySalt $u) ([byte[]]::new(16)) $FileKey CBC None
    $perms = Join-Bytes @([BitConverter]::GetBytes([int]$Permissions), [byte[]](0xFF, 0xFF, 0xFF, 0xFF), $latin1.GetBytes('Tadb'), (Get-RandomBytes 4))
    $permsEncrypted = Invoke-Aes $FileKey $null $perms ECB None
    return @{ U = $u; UE = $ue; O = $o; OE = $oe; Perms = $permsEncrypted }
}

function Protect-Data([byte[]] $FileKey, [byte[]] $Data, [string] $Method) {
    if ($Method -eq 'AESV4') {
        $nonce = Get-RandomBytes 12
        $ciphertext = [byte[]]::new($Data.Length); $tag = [byte[]]::new(16)
        $gcm = [System.Security.Cryptography.AesGcm]::new($FileKey, 16)
        try { $gcm.Encrypt($nonce, $Data, $ciphertext, $tag) } finally { $gcm.Dispose() }
        return , (Join-Bytes @($nonce, $ciphertext, $tag))
    }
    $iv = Get-RandomBytes 16
    return , (Join-Bytes @($iv, (Invoke-Aes $FileKey $iv $Data CBC PKCS7)))
}

function ConvertTo-Hex([byte[]] $Bytes) { '<' + [Convert]::ToHexString($Bytes) + '>' }

# A literal string: printable ASCII and bytes above 0x7F stay raw (']' included), other bytes use octal escapes.
function ConvertTo-LiteralString([byte[]] $Bytes) {
    $text = foreach ($b in $Bytes) { if (($b -ge 0x20 -and $b -ne 0x7F) -and $b -notin 0x28, 0x29, 0x5C) { [string][char]$b } else { '\' + [Convert]::ToString($b, 8).PadLeft(3, '0') } }
    '(' + (-join $text) + ')'
}

$padding = [byte[]](0x28, 0xBF, 0x4E, 0x5E, 0x4E, 0x75, 0x8A, 0x41, 0x64, 0x00, 0x4E, 0x56, 0xFF, 0xFA, 0x01, 0x08,
    0x2E, 0x2E, 0x00, 0xB6, 0xD0, 0x68, 0x3E, 0x80, 0x2F, 0x0C, 0xA9, 0xFE, 0x64, 0x53, 0x69, 0x7A)

function Invoke-Rc4([byte[]] $Key, [byte[]] $Data) {
    $s = [byte[]](0..255); $j = 0
    for ($i = 0; $i -lt 256; $i++) { $j = ($j + $s[$i] + $Key[$i % $Key.Length]) -band 255; $t = $s[$i]; $s[$i] = $s[$j]; $s[$j] = $t }
    $result = [byte[]]::new($Data.Length); $x = 0; $y = 0
    for ($i = 0; $i -lt $Data.Length; $i++) {
        $x = ($x + 1) -band 255; $y = ($y + $s[$x]) -band 255; $t = $s[$x]; $s[$x] = $s[$y]; $s[$y] = $t
        $result[$i] = $Data[$i] -bxor $s[($s[$x] + $s[$y]) -band 255]
    }
    return , $result
}

function Get-PaddedPassword([byte[]] $Password) { return , [byte[]](@($Password) + @($padding))[0..31] }
function Get-Md5([byte[]] $Data) { return , [System.Security.Cryptography.MD5]::HashData($Data) }
function Get-XorKey([byte[]] $Key, [int] $Value) { return , [byte[]]($Key | ForEach-Object { $_ -bxor $Value }) }

# ISO 32000-2 algorithms 2, 3 and 5 for revision 3 with a 40-bit key.
function New-Revision3Entries([byte[]] $UserPassword, [byte[]] $OwnerPassword, [int] $Permissions, [byte[]] $FileId, [int] $KeyLength) {
    $ownerKey = Get-Md5 (Get-PaddedPassword $OwnerPassword)
    # Algorithm 3 rehashes the whole digest each round; algorithm 2 below truncates every round.
    for ($i = 0; $i -lt 50; $i++) { $ownerKey = Get-Md5 $ownerKey }
    $ownerKey = [byte[]]$ownerKey[0..($KeyLength - 1)]
    $o = Invoke-Rc4 $ownerKey (Get-PaddedPassword $UserPassword)
    for ($i = 1; $i -le 19; $i++) { $o = Invoke-Rc4 (Get-XorKey $ownerKey $i) $o }
    $fileKey = Get-Md5 (Join-Bytes @((Get-PaddedPassword $UserPassword), $o, [BitConverter]::GetBytes([int]$Permissions), $FileId))
    for ($i = 0; $i -lt 50; $i++) { $fileKey = Get-Md5 ([byte[]]$fileKey[0..($KeyLength - 1)]) }
    $fileKey = [byte[]]$fileKey[0..($KeyLength - 1)]
    $u = Invoke-Rc4 $fileKey (Get-Md5 (Join-Bytes @($padding, $FileId)))
    for ($i = 1; $i -le 19; $i++) { $u = Invoke-Rc4 (Get-XorKey $fileKey $i) $u }
    return @{ O = $o; U = (Join-Bytes @($u, [byte[]]::new(16))); FileKey = $fileKey }
}

function Protect-Rc4Object([byte[]] $FileKey, [int] $ObjectNumber, [byte[]] $Data) {
    $objectKey = Get-Md5 (Join-Bytes @($FileKey, [byte[]](($ObjectNumber -band 255), (($ObjectNumber -shr 8) -band 255), (($ObjectNumber -shr 16) -band 255), 0, 0)))
    return , (Invoke-Rc4 ([byte[]]$objectKey[0..([Math]::Min(16, $FileKey.Length + 5) - 1)]) $Data)
}

function ConvertTo-PdfTextString([string] $Text) {
    return , [byte[]](@([System.Text.Encoding]::BigEndianUnicode.GetPreamble()) + @([System.Text.Encoding]::BigEndianUnicode.GetBytes($Text)))
}

function Write-Pdf([string] $Name, [string] $Header, [System.Collections.Generic.List[byte[]]] $Objects, [string] $FileId, [string] $TrailerExtra) {
    $output = [System.IO.MemoryStream]::new()
    $write = { param([byte[]] $bytes) $output.Write($bytes, 0, $bytes.Length) }
    & $write ($latin1.GetBytes("$Header`n%" + [char]0xE2 + [char]0xE3 + [char]0xCF + [char]0xD3 + "`n"))
    $offsets = @()
    for ($i = 0; $i -lt $Objects.Count; $i++) {
        $offsets += $output.Position
        & $write ($latin1.GetBytes("$($i + 1) 0 obj`n"))
        & $write $Objects[$i]
        & $write ($latin1.GetBytes("`nendobj`n"))
    }
    $xref = $output.Position
    $table = "xref`n0 $($Objects.Count + 1)`n0000000000 65535 f `n"
    foreach ($offset in $offsets) { $table += $offset.ToString('0000000000') + " 00000 n `n" }
    $table += "trailer`n<< /Size $($Objects.Count + 1) /Root 1 0 R /Info 6 0 R /Encrypt 7 0 R /ID [$FileId $FileId]$TrailerExtra >>`nstartxref`n$xref`n%%EOF`n"
    & $write ($latin1.GetBytes($table))
    [System.IO.File]::WriteAllBytes((Join-Path $OutputDirectory $Name), $output.ToArray())
}

function New-StreamObject([byte[]] $Data, [string] $Entries = '') {
    return , (Join-Bytes @($latin1.GetBytes("<< /Length $($Data.Length)$Entries >>`nstream`n"), $Data, $latin1.GetBytes("`nendstream")))
}

function New-PageObjects([byte[]] $ContentObject, [byte[]] $TitleString, [string] $CatalogExtra = '') {
    $objects = [System.Collections.Generic.List[byte[]]]::new()
    $objects.Add($latin1.GetBytes("<< /Type /Catalog /Pages 2 0 R$CatalogExtra >>"))
    $objects.Add($latin1.GetBytes('<< /Type /Pages /Kids [3 0 R] /Count 1 >>'))
    $objects.Add($latin1.GetBytes('<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 200] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>'))
    $objects.Add($ContentObject)
    $objects.Add($latin1.GetBytes('<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>'))
    $objects.Add($latin1.GetBytes("<< /Title $(ConvertTo-Hex $TitleString) >>"))
    return , $objects
}

function Get-PageContent([string] $Text) { return , $latin1.GetBytes("BT /F1 18 Tf 36 120 Td ($Text) Tj ET") }

# Revision 6 (AESV3) or revision 7 (AESV4). Modes: Document encrypts strings and streams; IdentityContentStream leaves the
# page content in clear text behind a /Crypt stream filter naming /Identity; DefaultMethodContentStream does the same through a
# named crypt filter that omits /CFM, which defaults to /None; EmbeddedFilesOnly encrypts only the embedded
# file (/StmF and /StrF /Identity, /EFF /StdCF, /AuthEvent /EFOpen).
function New-ModernFixture([string] $Name, [byte[]] $UserPassword, [byte[]] $OwnerPassword, [string] $Method, [string] $Text, [string] $Mode = 'Document') {
    $version = if ($Method -eq 'AESV4') { 6 } else { 5 }
    $revision = if ($Method -eq 'AESV4') { 7 } else { 6 }
    $permissions = -4
    $fileKey = Get-RandomBytes 32
    $entries = New-EncryptionEntries $UserPassword $OwnerPassword $fileKey $permissions
    $content = Get-PageContent $Text
    $title = ConvertTo-PdfTextString $Text
    $catalogExtra = ''
    $filters = '/CF << /StdCF << /AuthEvent /DocOpen /CFM /' + $Method + ' /Length 32 >> >> /StmF /StdCF /StrF /StdCF'
    switch ($Mode) {
        'Document' {
            $contentObject = New-StreamObject (Protect-Data $fileKey $content $Method)
            $title = Protect-Data $fileKey $title $Method
        }
        'IdentityContentStream' {
            $contentObject = New-StreamObject $content ' /Filter [/Crypt] /DecodeParms [<< /Type /CryptFilterDecodeParms /Name /Identity >>]'
            $title = Protect-Data $fileKey $title $Method
        }
        'DefaultMethodContentStream' {
            $contentObject = New-StreamObject $content ' /Filter [/Crypt] /DecodeParms [<< /Type /CryptFilterDecodeParms /Name /ClearCF >>]'
            $title = Protect-Data $fileKey $title $Method
            $filters = '/CF << /StdCF << /AuthEvent /DocOpen /CFM /' + $Method + ' /Length 32 >> /ClearCF << /AuthEvent /DocOpen >> >> /StmF /StdCF /StrF /StdCF'
        }
        'EmbeddedFilesOnly' {
            $contentObject = New-StreamObject $content
            $catalogExtra = ' /Names << /EmbeddedFiles << /Names [(note.txt) 8 0 R] >> >>'
            $filters = '/CF << /StdCF << /AuthEvent /EFOpen /CFM /' + $Method + ' /Length 32 >> >> /StmF /Identity /StrF /Identity /EFF /StdCF'
        }
    }

    $objects = New-PageObjects $contentObject $title $catalogExtra
    $objects.Add($latin1.GetBytes("<< /Filter /Standard /V $version /R $revision /Length 256 /P $permissions $filters /O $(ConvertTo-Hex $entries.O) /U $(ConvertTo-Hex $entries.U) /OE $(ConvertTo-Hex $entries.OE) /UE $(ConvertTo-Hex $entries.UE) /Perms $(ConvertTo-Hex $entries.Perms) >>"))
    if ($Mode -eq 'EmbeddedFilesOnly') {
        $objects.Add($latin1.GetBytes('<< /Type /Filespec /F (note.txt) /UF (note.txt) /EF << /F 9 0 R >> >>'))
        $objects.Add((New-StreamObject (Protect-Data $fileKey $latin1.GetBytes('Attachment secret') $Method) ' /Type /EmbeddedFile'))
    }

    $header = if ($revision -eq 7) { '%PDF-2.0' } else { '%PDF-1.7' }
    Write-Pdf $Name $header $objects (ConvertTo-Hex (Get-RandomBytes 16)) ''
}

# Revision 3 without /Length: /V 1 with the 40-bit default, or /V 2 from a producer that wrote a 128-bit key anyway.
function New-Revision3Fixture([string] $Name, [byte[]] $UserPassword, [byte[]] $OwnerPassword, [string] $Text, [int] $Version = 1, [int] $KeyLength = 5, [switch] $LiteralIdWithBracket) {
    $permissions = -1852
    $fileIdBytes = Get-RandomBytes 16
    if ($LiteralIdWithBracket) { $fileIdBytes[1] = 0x5D; $fileIdBytes[2] = 0xE9 }
    $entries = New-Revision3Entries $UserPassword $OwnerPassword $permissions $fileIdBytes $KeyLength
    $objects = New-PageObjects (New-StreamObject (Protect-Rc4Object $entries.FileKey 4 (Get-PageContent $Text))) (Protect-Rc4Object $entries.FileKey 6 (ConvertTo-PdfTextString $Text))
    $objects.Add($latin1.GetBytes("<< /Filter /Standard /V $Version /R 3 /P $permissions /O $(ConvertTo-Hex $entries.O) /U $(ConvertTo-Hex $entries.U) >>"))
    $fileId = if ($LiteralIdWithBracket) { ConvertTo-LiteralString $fileIdBytes } else { ConvertTo-Hex $fileIdBytes }
    Write-Pdf $Name '%PDF-1.4' $objects $fileId ''
}

$utf8 = [System.Text.Encoding]::UTF8
# RFC 4013 maps U+00AA to "a" (NFKC) and removes U+00AD (table B.1): the stored password is "SaSLprep".
New-ModernFixture 'r6-saslprep.pdf' $utf8.GetBytes('SaSLprep') $utf8.GetBytes('owner') 'AESV3' 'SASLprep R6'
# A producer that skipped the SASLprep mapping stored NFKC("pass" U+00AD "word") with the soft hyphen kept.
New-ModernFixture 'r6-unprepared-password.pdf' $utf8.GetBytes("pass$([char]0x00AD)word") $utf8.GetBytes('owner') 'AESV3' 'Unprepared R6'
New-ModernFixture 'r6-identity-crypt-filter-stream.pdf' $utf8.GetBytes('open') $utf8.GetBytes('owner') 'AESV3' 'Identity stream R6' 'IdentityContentStream'
New-ModernFixture 'r6-crypt-filter-without-method.pdf' $utf8.GetBytes('open') $utf8.GetBytes('owner') 'AESV3' 'Default method R6' 'DefaultMethodContentStream'
New-ModernFixture 'r6-embedded-files-only.pdf' $utf8.GetBytes('open') $utf8.GetBytes('owner') 'AESV3' 'Attachments only R6' 'EmbeddedFilesOnly'
New-ModernFixture 'r7-aes-gcm.pdf' $utf8.GetBytes('open') $utf8.GetBytes('owner') 'AESV4' 'AES-GCM R7'
New-Revision3Fixture 'r3-rc4-40-default-length.pdf' $utf8.GetBytes('open') $utf8.GetBytes('owner') 'RC4 40 R3'
New-Revision3Fixture 'r3-rc4-128-without-length.pdf' $utf8.GetBytes('open') $utf8.GetBytes('owner') 'RC4 128 R3' 2 16
# The permanent identifier is a literal string containing ']', which rewrites must not cut short.
New-Revision3Fixture 'r3-literal-id-with-bracket.pdf' $utf8.GetBytes('open') $utf8.GetBytes('owner') 'Literal ID R3' -LiteralIdWithBracket
