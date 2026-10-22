<#
.SYNOPSIS
Saves and verifies the exact tools used by the PowerShell CI lanes.
.DESCRIPTION
Uses PSResourceGet on PowerShell 7, avoiding PowerShellGet's NuGet.exe command
line. Windows PowerShell uses its supported PowerShellGet provider. The cache
is local to the runner; GITHUB_ENV exposes it to subsequent test processes.
#>
$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
$moduleRoot = Join-Path $env:RUNNER_TEMP 'officeimo-ci-modules'
$null = New-Item -ItemType Directory -Force -Path $moduleRoot
$modules = @{
    Pester = '6.1.0'
    PSWriteColor = '1.0.3'
}
if ($env:OFFICEIMO_CI_PACKAGING -eq 'true') {
    $modules.PSPublishModule = '3.0.170'
}

if ($PSVersionTable.PSEdition -eq 'Core') {
    Import-Module Microsoft.PowerShell.PSResourceGet -ErrorAction Stop
    if (-not (Get-PSResourceRepository -Name PSGallery -ErrorAction SilentlyContinue)) {
        Register-PSResourceRepository -PSGallery
    }
} else {
    try {
        $repository = Get-PSRepository -Name PSGallery -ErrorAction Stop
        if (-not $repository.SourceLocation) { throw 'PSGallery has no source URL.' }
    } catch {
        if ($_.FullyQualifiedErrorId -notmatch '^(SourceNotFound|RepositoryNotFound),') { throw }
        Register-PSRepository -Default
    }
    $null = Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force
}

foreach ($name in $modules.Keys) {
    $version = $modules[$name]
    $manifest = Join-Path $moduleRoot "$name/$version/$name.psd1"
    if (-not (Test-Path -LiteralPath $manifest)) {
        if ($PSVersionTable.PSEdition -eq 'Core') {
            Save-PSResource -Name $name -Version $version -Repository PSGallery -Path $moduleRoot -TrustRepository -Quiet
        } else {
            Save-Module -Name $name -RequiredVersion $version -Repository PSGallery -Path $moduleRoot -Force
        }
    }
    $metadata = Import-PowerShellDataFile -LiteralPath $manifest
    if ([string]$metadata.ModuleVersion -ne $version) { throw "Unexpected $name version in $manifest" }
}

$separator = [IO.Path]::PathSeparator
$env:PSModulePath = $moduleRoot + $separator + $env:PSModulePath
"PSModulePath=$env:PSModulePath" | Out-File -FilePath $env:GITHUB_ENV -Append -Encoding utf8
"PESTER_VERSION=6.1.0" | Out-File -FilePath $env:GITHUB_ENV -Append -Encoding utf8
Import-Module Pester -RequiredVersion 6.1.0 -Force
if ((Get-Module Pester).Version -ne [version]'6.1.0') { throw 'Incorrect Pester loaded.' }
