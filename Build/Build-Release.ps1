<#
.SYNOPSIS
Builds or publishes OfficeIMO and PSWriteOffice as one coordinated release.
.DESCRIPTION
The native PowerForge pipeline builds packages before the module, synchronizes
their version, and publishes NuGet before PowerShell Gallery and GitHub.
ResumeIncompleteRelease reuses the shared publisher's release checkpoint.
Publish first checks that every package declares and packs the notices for its
third-party material (Build/validate_third_party_notices.py), and stops before
anything is built or published when it does not.
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
    [switch] $JsonOnly,
    [string] $JsonPath,
    [switch] $ResumeIncompleteRelease,
    [switch] $PassThru
)

if ($RunMode -eq 'Publish') {
    $python = Get-Command python3, python -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $python) {
        throw 'Python is required to check third-party notices before publishing.'
    }
    & $python.Source "$PSScriptRoot/validate_third_party_notices.py"
    if ($LASTEXITCODE -ne 0) {
        throw 'Packages do not declare or pack their third-party notices; nothing was published.'
    }
}

& "$PSScriptRoot/Build-Module.ps1" @PSBoundParameters
