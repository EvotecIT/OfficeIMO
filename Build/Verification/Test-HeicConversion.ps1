<#
.SYNOPSIS
Verifies HEIC to PDF conversion through the imported PSWriteOffice module.
.DESCRIPTION
Imports the supplied module, converts a fixture through the normal document conversion command,
and verifies the PDF artifact and successful workflow result. Requires PowerShell 7.4 or newer
and a supported system codec. Existing output is not overwritten.
.PARAMETER ModulePath
Path to the built PSWriteOffice module manifest.
.PARAMETER InputPath
HEIC, HEIF or HIF source file.
.PARAMETER OutputPath
New PDF destination path in a task-owned output directory.
.EXAMPLE
./Build/Verification/Test-HeicConversion.ps1 -ModulePath ./PowerShell/PSWriteOffice/PSWriteOffice.psd1 -InputPath ./OfficeIMO.TestAssets/Images/Heic/basic-8bit-srgb.heic -OutputPath ./Artefacts/HeicSmoke/powershell.pdf
.OUTPUTS
The successful workflow result.
#>
#requires -Version 7.4
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string] $ModulePath,

    [Parameter(Mandatory)]
    [string] $InputPath,

    [Parameter(Mandatory)]
    [string] $OutputPath
)

$ErrorActionPreference = 'Stop'
Import-Module $ModulePath -ErrorAction Stop
$result = ConvertTo-OfficeDocument -InputPath $InputPath -OutputPath $OutputPath -ErrorAction Stop
if (-not (Test-Path -LiteralPath $OutputPath -PathType Leaf)) {
    throw 'The conversion did not create the PDF.'
}
$bytes = [IO.File]::ReadAllBytes($OutputPath)
if ($bytes.Length -lt 5 -or [Text.Encoding]::ASCII.GetString($bytes, 0, 5) -ne '%PDF-') {
    throw 'The conversion output is not a PDF.'
}
if ($null -eq $result -or -not $result.PSObject.Properties['Succeeded'] -or -not $result.Succeeded) {
    throw 'The workflow did not report success.'
}
$result
