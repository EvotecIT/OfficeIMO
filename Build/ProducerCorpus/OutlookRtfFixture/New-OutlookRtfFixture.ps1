[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string] $OutputPath
)

$ErrorActionPreference = 'Stop'
$outlook = $null
$mail = $null

try {
    $outlook = New-Object -ComObject Outlook.Application
    $mail = $outlook.CreateItem(0)
    $mail.Subject = 'OfficeIMO RTF producer fixture'
    $mail.BodyFormat = 2
    $mail.HTMLBody = @'
<html><body><h1>Outlook MAPI HTML fixture</h1><p>Zażółć gęślą jaźń</p><ul><li>First</li><li>Second</li></ul><table><tr><th>Key</th><th>Value</th></tr><tr><td>Bandage</td><td>Ready</td></tr></table></body></html>
'@

    [byte[]] $bytes = $mail.RTFBody
    if ($bytes.Length -eq 0) {
        throw 'Outlook did not generate an RTF body.'
    }

    $resolvedOutput = [System.IO.Path]::GetFullPath($OutputPath)
    $directory = [System.IO.Path]::GetDirectoryName($resolvedOutput)
    [System.IO.Directory]::CreateDirectory($directory) | Out-Null
    [System.IO.File]::WriteAllBytes($resolvedOutput, $bytes)
    $sha256 = (Get-FileHash -LiteralPath $resolvedOutput -Algorithm SHA256).Hash.ToLowerInvariant()

    [pscustomobject]@{
        OutlookVersion = $outlook.Version
        OutputPath = $resolvedOutput
        Bytes = $bytes.Length
        Sha256 = $sha256
    }
} finally {
    if ($mail) {
        [void] [Runtime.InteropServices.Marshal]::FinalReleaseComObject($mail)
    }
    if ($outlook) {
        [void] [Runtime.InteropServices.Marshal]::FinalReleaseComObject($outlook)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
