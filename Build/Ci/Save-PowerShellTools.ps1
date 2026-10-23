<#
.SYNOPSIS
Saves the repository's pinned, dependency-free PowerShell CI tools.
.DESCRIPTION
Downloads exact Gallery archives without a metadata lookup or NuGet provider.
Checks their SHA-256 and manifest before exposing them to module discovery.
This is a bootstrap for the three tools in powershell-tools.json, not a general
module dependency installer. Version updates must recheck package dependencies.
.PARAMETER ModuleRoot
Module search directory receiving the standard Name/Version module folders.
.PARAMETER Tool
Pinned tools to save. Defaults to the two tools used by the Pester lanes.
.EXAMPLE
./Build/Ci/Save-PowerShellTools.ps1 -ModuleRoot ./Artefacts/TestModules
#>
param(
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string] $ModuleRoot,
    [ValidateSet('Pester', 'PSWriteColor', 'PSPublishModule')]
    [string[]] $Tool = @('Pester', 'PSWriteColor')
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
$pins = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'powershell-tools.json') -Raw | ConvertFrom-Json
$null = New-Item -ItemType Directory -Force -Path $ModuleRoot

foreach ($name in $Tool) {
    $pin = $pins.$name
    $destination = Join-Path $ModuleRoot "$name/$($pin.version)"
    $manifest = Join-Path $destination "$name.psd1"
    if (-not (Test-Path -LiteralPath $manifest)) {
        # A unique staging directory keeps interrupted downloads out of the cache.
        $staging = Join-Path ([IO.Path]::GetTempPath()) ('officeimo-ci-tools-' + [guid]::NewGuid().ToString('N'))
        $null = New-Item -ItemType Directory -Path $staging
        try {
            $archive = Join-Path $staging 'package.zip'
            $uri = "https://cdn.powershellgallery.com/packages/$($name.ToLowerInvariant()).$($pin.version).nupkg"
            for ($attempt = 1; $attempt -le 3; $attempt++) {
                try {
                    Invoke-WebRequest -Uri $uri -OutFile $archive -UseBasicParsing -TimeoutSec 90
                    break
                } catch {
                    if ($attempt -eq 3) { throw }
                    Write-Warning "Download of $name failed; retrying ($attempt/3)."
                    Start-Sleep -Seconds 2
                }
            }
            if ((Get-FileHash -LiteralPath $archive -Algorithm SHA256).Hash -ne $pin.sha256) {
                throw "The $name archive does not match powershell-tools.json."
            }
            $expanded = Join-Path $staging 'module'
            Expand-Archive -LiteralPath $archive -DestinationPath $expanded
            $metadata = Import-PowerShellDataFile -LiteralPath (Join-Path $expanded "$name.psd1")
            if ([string]$metadata.ModuleVersion -ne $pin.version -or $metadata.RequiredModules) {
                throw "Unexpected version or dependencies in the pinned $name package."
            }
            $null = New-Item -ItemType Directory -Force -Path (Split-Path -Parent $destination)
            Move-Item -LiteralPath $expanded -Destination $destination
        } finally {
            # This directory was created above and contains only the verified archive.
            Remove-Item -LiteralPath $staging -Recurse -Force
        }
    }
    $metadata = Import-PowerShellDataFile -LiteralPath $manifest
    if ([string]$metadata.ModuleVersion -ne $pin.version -or $metadata.RequiredModules) {
        throw "Unexpected version or dependencies in the cached $name manifest."
    }
    Write-Host "Verified $name $($pin.version) in $destination"
}
