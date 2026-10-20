<#
.SYNOPSIS
Builds or publishes OfficeIMO and PSWriteOffice as one coordinated release.
.DESCRIPTION
The native PowerForge pipeline builds packages before the module, synchronizes
their version, and publishes NuGet before PowerShell Gallery and GitHub.
ResumeIncompleteRelease reuses the shared publisher's release checkpoint.
.EXAMPLE
./Build/Build-Release.ps1
.EXAMPLE
./Build/Build-Release.ps1 -RunMode Publish -SignModule:$true
#>
param(
    [ValidateSet('Documentation', 'Manifest', 'Build', 'Publish')]
    [string] $RunMode = 'Build',
    [bool] $SignModule = $false,
    [string] $StagingPath,
    [string] $PowerShellGalleryApiKeyFile = $env:PSGALLERY_API_KEY_FILE,
    [string] $GitHubTokenFile = $env:GITHUB_TOKEN_FILE,
    [switch] $JsonOnly,
    [string] $JsonPath,
    [switch] $ResumeIncompleteRelease,
    [switch] $PassThru
)

& "$PSScriptRoot/Build-Module.ps1" @PSBoundParameters
