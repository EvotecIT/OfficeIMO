param(
    [string] $OutputDirectory = (Join-Path $PSScriptRoot '..\..\Artefacts\Examples\Visio'),
    [string] $StencilPackagePath,
    [switch] $SkipPackageExample
)

$ErrorActionPreference = 'Stop'

$results = @()
$results += & (Join-Path $PSScriptRoot 'Example-Visio-StencilFlow.ps1') -OutputDirectory $OutputDirectory
$results += & (Join-Path $PSScriptRoot 'Example-Visio-ArchitectureMap.ps1') -OutputDirectory $OutputDirectory
$results += & (Join-Path $PSScriptRoot 'Example-Visio-NetworkTopology.ps1') -OutputDirectory $OutputDirectory

if (-not $SkipPackageExample) {
    $parameters = @{ OutputDirectory = $OutputDirectory }
    if ($StencilPackagePath) {
        $parameters.StencilPackagePath = $StencilPackagePath
    }

    $results += & (Join-Path $PSScriptRoot 'Example-Visio-PackageStencil.ps1') @parameters
}

$results | Format-Table -AutoSize
$results
