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
    $nodeExe = (Get-Command node -ErrorAction Stop).Source
    $vsceBin = Join-Path $repoRoot 'node_modules/.bin/vsce.cmd'
    if (-not (Test-Path -LiteralPath $vsceBin)) {
        $vsceBin = Join-Path $repoRoot 'node_modules/.bin/vsce'
    }
    $vsceMain = Join-Path $repoRoot 'node_modules/@vscode/vsce/out/main.js'
    $esbuildPath = Join-Path $repoRoot 'node_modules/esbuild'

    $nodeModules = Join-Path $repoRoot 'node_modules'
    $installNeeded = (-not (Test-Path -LiteralPath $nodeModules)) -or
        (-not (Test-Path -LiteralPath $esbuildPath)) -or
        (-not (Test-Path -LiteralPath $vsceBin) -and -not (Test-Path -LiteralPath $vsceMain))

    if ($installNeeded) {
        Write-Host "Installing dev dependencies (typescript/esbuild/vsce)..." -ForegroundColor Yellow
        npm install --include=dev
    }

    if (-not (Test-Path -LiteralPath $esbuildPath)) {
        throw "esbuild not found at $esbuildPath. Run npm install."
    }
    if (-not (Test-Path -LiteralPath $vsceBin) -and -not (Test-Path -LiteralPath $vsceMain)) {
        throw "VSCE not found in node_modules. Run npm install."
    }

    dotnet build ..\OfficeIMO.Markup.Cli\OfficeIMO.Markup.Cli.csproj -c Release --framework net8.0 --no-restore -m:1 -p:BuildInParallel=false -p:UseSharedCompilation=false --nologo --verbosity minimal

    $cliOutput = Resolve-Path -LiteralPath (Join-Path $repoRoot '..\OfficeIMO.Markup.Cli\bin\Release\net8.0')
    $bundledCli = Join-Path $repoRoot 'tools\OfficeIMO.Markup.Cli'
    if (Test-Path -LiteralPath $bundledCli) {
        Remove-Item -LiteralPath $bundledCli -Recurse -Force
    }

    New-Item -ItemType Directory -Path $bundledCli | Out-Null
    Copy-Item -Path (Join-Path $cliOutput '*') -Destination $bundledCli -Recurse -Force

    Write-Host "Building OfficeIMO Markup VSIX..." -ForegroundColor Cyan
    npm run compile
    if (Test-Path -LiteralPath $vsceBin) {
        & $vsceBin package --allow-missing-repository
    } else {
        & $nodeExe $vsceMain package --allow-missing-repository
    }

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
