<#
.SYNOPSIS
Compiles affected library consumers and the build lane's executable contracts.
.DESCRIPTION
Uses the selector's conservative project closure instead of building every test,
example and desktop application on both operating systems for an ordinary PR.
Master, nightly and full-ci retain the original complete solution build.
.PARAMETER ProjectsJson
JSON array of repository-relative affected project paths from select.py.
.PARAMETER ArtifactsPath
Optional directory for local qualification output outside the checkout.
#>
param(
    [Parameter(Mandatory)][string] $ProjectsJson,
    [string] $ArtifactsPath
)
$ErrorActionPreference = 'Stop'
$PSNativeCommandUseErrorActionPreference = $true
$solutionProjects = @(dotnet sln OfficeIMO.sln list | Where-Object { $_ -match '\.csproj$' } | ForEach-Object {
    $_.Replace('\', '/')
})
$projects = @($ProjectsJson | ConvertFrom-Json | Where-Object {
    $_ -in $solutionProjects -and $_ -like 'OfficeIMO.*/*.csproj' -and $_ -notmatch '(Tests|Benchmarks|Examples|Example|Studio|SamplePlugin)'
})
$projects += @(
    'OfficeIMO.Provenance.C2pa.Tests/OfficeIMO.Provenance.C2pa.Tests.csproj'
    'OfficeIMO.Invoicing.Tests/OfficeIMO.Invoicing.Tests.csproj'
    'OfficeIMO.Invoicing.Pdf.Tests/OfficeIMO.Invoicing.Pdf.Tests.csproj'
)
$projectPaths = @($projects | Sort-Object -Unique | ForEach-Object { (Resolve-Path -LiteralPath $_).Path })
$solutionRoot = Join-Path $env:RUNNER_TEMP 'officeimo-pr-build'
dotnet new sln --name OfficeIMO.Pr --output $solutionRoot --format sln --force
$solution = Join-Path $solutionRoot 'OfficeIMO.Pr.sln'
dotnet sln $solution add @projectPaths
$artifactArguments = @()
if ($ArtifactsPath) { $artifactArguments = @('--artifacts-path', $ArtifactsPath) }
# One graph build shares dependency compilation across all affected consumers.
# Preserve each project's declared targets, including netstandard generators.
dotnet build $solution -c $env:BUILD_CONFIGURATION -m:2 @artifactArguments
