param(
    [switch]$Force
)

$ErrorActionPreference = 'Stop'

$repoRoot = Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')
$packagePath = Join-Path $repoRoot 'package.json'
if (-not (Test-Path -LiteralPath $packagePath)) {
    throw "package.json not found at $packagePath"
}

Push-Location $repoRoot
try {
    $packageScript = Join-Path $repoRoot 'scripts/package-vsix.ps1'
    if (-not (Test-Path -LiteralPath $packageScript)) {
        throw "Package script not found at $packageScript"
    }

    & $packageScript -OutputDirectory $repoRoot

    $package = Get-Content -LiteralPath $packagePath -Raw | ConvertFrom-Json
    $vsixName = "{0}-{1}.vsix" -f $package.name, $package.version
    $vsixPath = Join-Path $repoRoot $vsixName
    if (-not (Test-Path -LiteralPath $vsixPath)) {
        throw "VSIX not found at $vsixPath"
    }

    $forceFlag = if ($Force) { "--force" } else { "" }
    Write-Host "Installing into VS Code Insiders..." -ForegroundColor Cyan
    & code-insiders --install-extension $vsixPath $forceFlag

    Write-Host "Installed. Reload VS Code Insiders to activate the update." -ForegroundColor Green
} finally {
    Pop-Location
}
