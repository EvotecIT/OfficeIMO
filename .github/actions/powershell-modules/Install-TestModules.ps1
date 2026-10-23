<#
.SYNOPSIS
Saves and verifies the exact tools used by the PowerShell CI lanes.
.DESCRIPTION
Uses the repository's pinned archive bootstrap on both PowerShell hosts. The
cache is local to the runner; GITHUB_ENV exposes it to subsequent test processes.
#>
$ErrorActionPreference = 'Stop'
$moduleRoot = Join-Path $env:RUNNER_TEMP 'officeimo-ci-modules'
$tools = @('Pester', 'PSWriteColor')
if ($env:OFFICEIMO_CI_PACKAGING -eq 'true') {
    $tools += 'PSPublishModule'
}
& (Join-Path $PSScriptRoot '../../../Build/Ci/Save-PowerShellTools.ps1') -ModuleRoot $moduleRoot -Tool $tools

$separator = [IO.Path]::PathSeparator
$env:PSModulePath = $moduleRoot + $separator + $env:PSModulePath
"PSModulePath=$env:PSModulePath" | Out-File -FilePath $env:GITHUB_ENV -Append -Encoding utf8
"PESTER_VERSION=6.1.0" | Out-File -FilePath $env:GITHUB_ENV -Append -Encoding utf8
Import-Module Pester -RequiredVersion 6.1.0 -Force
if ((Get-Module Pester).Version -ne [version]'6.1.0') { throw 'Incorrect Pester loaded.' }
