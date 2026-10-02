param(
    [string] $PowerForgeCliPath = $env:POWERFORGE_CLI,
    [switch] $Validate,
    [switch] $Plan
)

# PowerForge owns publishing, nested-code signing, and certificate/team validation.
$configPath = Join-Path $PSScriptRoot 'Apple/powerforge.development.json'
$arguments = @('dotnet', 'publish', '--config', $configPath)
if ($Validate) { $arguments += '--validate' }
if ($Plan) { $arguments += '--plan' }

if ($PowerForgeCliPath) {
    & dotnet $PowerForgeCliPath @arguments
} else {
    & powerforge @arguments
}
if ($LASTEXITCODE -ne 0) { throw "Studio development build failed with exit code $LASTEXITCODE." }
