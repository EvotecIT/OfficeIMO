param(
    [Parameter(Mandatory)][string] $PackagePath,
    [string] $BundleDirectory = "$PSScriptRoot/../../OfficeIMO.JavaScript/bundles"
)
$ErrorActionPreference = 'Stop'
# Preserve the signed native-release check for operators using this entrypoint.
& dotnet nuget verify $PackagePath --all --certificate-fingerprint D13C03BFEEC0CBF7FD0E0E7FFAF3F3D2076E62B1AF64665E806917B9191CFFDE | Out-Host
if ($LASTEXITCODE -ne 0) { throw 'The released OfficeIMO.JavaScript.Assets package signature is not valid.' }
& "$PSScriptRoot/Test-NpmAssetParity.ps1" -PackagePath $PackagePath -BundleDirectory $BundleDirectory
