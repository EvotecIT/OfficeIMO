param(
    [ValidateSet('Studio.Windows', 'Studio.macOS', 'Studio.Linux')]
    [string] $Target,
    [string[]] $Runtime,
    [switch] $Validate,
    [switch] $Plan
)

if ($Runtime -and -not $Target) {
    throw 'Use -Target when narrowing Studio runtimes, or omit both parameters for the configured release matrix.'
}

if ($Runtime) {
    $allowedRuntimes = @{
        'Studio.Windows' = @('win-x64', 'win-arm64')
        'Studio.macOS' = @('osx-x64', 'osx-arm64')
        'Studio.Linux' = @('linux-x64', 'linux-arm64')
    }
    $invalidRuntimes = @($Runtime | Where-Object { $_ -notin $allowedRuntimes[$Target] })
    if ($invalidRuntimes.Count -gt 0) {
        throw "Runtime '$($invalidRuntimes -join ', ')' is not valid for target '$Target'."
    }
}

$repositoryRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$configPath = Join-Path $PSScriptRoot 'powerforge.dotnetpublish.json'
$previousNuGetPackages = $env:NUGET_PACKAGES
$env:NUGET_PACKAGES = Join-Path $repositoryRoot '.nuget/packages'

Import-Module PSPublishModule -MinimumVersion 3.0.131 -Force

$parameters = @{
    ConfigPath = $configPath
    ExitCode = $true
    ErrorAction = 'Stop'
}
if ($Target) { $parameters.Target = $Target }
if ($Runtime) { $parameters.Runtimes = $Runtime }
if ($Validate) { $parameters.Validate = $true }
if ($Plan) { $parameters.Plan = $true }

Push-Location $repositoryRoot
try {
    Invoke-DotNetPublish @parameters
} finally {
    Pop-Location
    $env:NUGET_PACKAGES = $previousNuGetPackages
}
