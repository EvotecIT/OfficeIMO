param(
    [string[]] $Target,
    [string[]] $Runtime,
    [switch] $Validate,
    [switch] $Plan
)

$repositoryRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$configPath = Join-Path $PSScriptRoot 'powerforge.dotnetpublish.json'

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
}
