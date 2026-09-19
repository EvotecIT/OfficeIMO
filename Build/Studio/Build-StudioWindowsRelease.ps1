param(
    [switch] $Validate,
    [switch] $Plan,
    [switch] $Publish,
    [switch] $SubmitWinget
)

if ($SubmitWinget -and -not $Publish) {
    throw 'WinGet submission requires the exact signed GitHub release to be published first.'
}

$repositoryRoot = (Resolve-Path (Join-Path $PSScriptRoot '../..')).Path
$previousNuGetPackages = $env:NUGET_PACKAGES
$env:NUGET_PACKAGES = Join-Path $repositoryRoot '.nuget/packages'

Import-Module PSPublishModule -MinimumVersion 3.0.145 -Force

$parameters = @{
    ConfigPath = Join-Path $PSScriptRoot 'powerforge.windows-release.json'
    ToolsOnly = $true
    Target = @('Studio.Windows')
    Runtimes = @('win-x64', 'win-arm64')
    ExitCode = $true
    ErrorAction = 'Stop'
}
if ($Validate) { $parameters.Validate = $true }
if ($Plan) { $parameters.Plan = $true }
if ($Publish) { $parameters.PublishProjectGitHub = $true }
if ($SubmitWinget) { $parameters.SubmitWinget = $true }

Push-Location $repositoryRoot
try {
    Invoke-PowerForgeRelease @parameters
} finally {
    Pop-Location
    $env:NUGET_PACKAGES = $previousNuGetPackages
}
