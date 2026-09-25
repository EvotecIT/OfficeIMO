param(
    [switch] $Validate,
    [switch] $Plan
)

$repositoryRoot = (Resolve-Path (Join-Path $PSScriptRoot '../..')).Path
$previousNuGetPackages = $env:NUGET_PACKAGES

$env:NUGET_PACKAGES = Join-Path $repositoryRoot '.nuget/packages'

$parameters = @{
    ConfigPath = Join-Path $PSScriptRoot 'powerforge.linux-release.json'
    ToolsOnly = $true
    Target = @('Studio.Linux')
    Runtimes = @('linux-x64', 'linux-arm64')
    ExitCode = $true
    ErrorAction = 'Stop'
}
if ($Validate) { $parameters.Validate = $true }
if ($Plan) { $parameters.Plan = $true }

try {
    Import-Module PSPublishModule -MinimumVersion 3.0.151 -Force -ErrorAction Stop
    Push-Location $repositoryRoot -ErrorAction Stop
    try {
        Invoke-PowerForgeRelease @parameters
    } finally {
        Pop-Location
    }
} finally {
    $env:NUGET_PACKAGES = $previousNuGetPackages
}
