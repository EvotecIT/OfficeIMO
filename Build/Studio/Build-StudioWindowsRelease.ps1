param(
    [switch] $Validate,
    [switch] $Plan,
    [switch] $Publish
)

$repositoryRoot = (Resolve-Path (Join-Path $PSScriptRoot '../..')).Path
$previousNuGetPackages = $env:NUGET_PACKAGES
$env:NUGET_PACKAGES = Join-Path $repositoryRoot '.nuget/packages'

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

try {
    Import-Module PSPublishModule -MinimumVersion 3.0.145 -Force -ErrorAction Stop
    Push-Location $repositoryRoot -ErrorAction Stop
    try {
        Invoke-PowerForgeRelease @parameters
    } finally {
        Pop-Location
    }
} finally {
    $env:NUGET_PACKAGES = $previousNuGetPackages
}
