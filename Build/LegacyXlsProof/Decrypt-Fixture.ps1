# Decrypt the test's randomized XLSX output for package comparison.
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string] $CoreAssembly,
    [Parameter(Mandatory)][string] $Package
)

$ErrorActionPreference = 'Stop'
$assembly = [System.Reflection.Assembly]::LoadFrom((Resolve-Path -LiteralPath $CoreAssembly).ProviderPath)
$encryption = $assembly.GetType('OfficeIMO.Core.Internal.OfficeEncryption', $true)
$decrypt = $encryption.GetMethod('DecryptPackage', [type[]] @([byte[]], [string]))
$arguments = [object[]]::new(2)
$arguments[0] = [IO.File]::ReadAllBytes($Package)
# This password belongs to LegacyXls_NormalLoad_SaveEncryptedPathProducesEncryptedXlsx.
$arguments[1] = 'legacy-xls-secret'
$plaintext = $decrypt.Invoke($null, $arguments)
[IO.File]::WriteAllBytes([IO.Path]::ChangeExtension($Package, 'plaintext.xlsx'), $plaintext)
