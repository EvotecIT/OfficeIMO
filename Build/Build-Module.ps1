<#
.SYNOPSIS
Builds PSWriteOffice and its OfficeIMO packages through the shared release pipeline.
.DESCRIPTION
Build is unsigned by default and does not install the module. Publish coordinates
NuGet, PowerShell Gallery and one OfficeIMO GitHub release. Use IncludePackages
only to skip package creation during local module iteration.
.EXAMPLE
./Build/Build-Module.ps1 -RunMode Build -IncludePackages:$false
.EXAMPLE
./Build/Build-Module.ps1 -RunMode Publish -SignModule:$true
#>
param(
    [ValidateSet('Documentation', 'Manifest', 'Build', 'Publish')]
    [string] $RunMode = 'Build',
    [bool] $SignModule = $false,
    [bool] $IncludePackages = $true,
    [string] $StagingPath,
    [string] $PowerShellGalleryApiKeyFile = $env:PSGALLERY_API_KEY_FILE,
    [string] $GitHubTokenFile = $env:GITHUB_TOKEN_FILE,
    [switch] $JsonOnly,
    [string] $JsonPath,
    [switch] $ResumeIncompleteRelease,
    [switch] $PassThru
)

$ErrorActionPreference = 'Stop'
if ($RunMode -eq 'Publish' -and (-not $IncludePackages -or -not $SignModule)) {
    throw 'Publish requires the coordinated package build and signing. Use Build for unsigned local validation.'
}
& "$PSScriptRoot/../PowerShell/PSWriteOffice/Build/Build-Module.ps1" @PSBoundParameters -SkipInstall
