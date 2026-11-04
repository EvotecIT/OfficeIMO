param(
    [string] $PowerForgeCliPath = $env:POWERFORGE_CLI,
    [switch] $Validate,
    [switch] $Plan
)

$ErrorActionPreference = 'Stop'

# PowerForge owns publishing, nested-code signing, and certificate/team validation.
$configPath = Join-Path $PSScriptRoot 'Apple/powerforge.development.json'
$arguments = @('dotnet', 'publish', '--config', $configPath)
if ($Validate) { $arguments += '--validate' }
if ($Plan) { $arguments += '--plan' }

if ($PowerForgeCliPath) {
    & dotnet $PowerForgeCliPath @arguments
} else {
    $powerforge = Get-Command -Name powerforge -CommandType Application -ErrorAction SilentlyContinue
    if (-not $powerforge) {
        throw 'PowerForge CLI was not found. Install powerforge or pass -PowerForgeCliPath with the path to PowerForge.Cli.dll (also supported through POWERFORGE_CLI).'
    }
    & $powerforge.Source @arguments
}
if ($LASTEXITCODE -ne 0) { throw "Studio development build failed with exit code $LASTEXITCODE." }
