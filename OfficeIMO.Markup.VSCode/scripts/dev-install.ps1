param(
    [switch]$Insiders,
    [switch]$Stable,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'

$repoRoot = Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')
$packagePath = Join-Path $repoRoot 'package.json'
if (-not (Test-Path -LiteralPath $packagePath)) {
    throw "package.json not found at $packagePath"
}

$package = Get-Content -LiteralPath $packagePath -Raw | ConvertFrom-Json
$publisher = $package.publisher
$name = $package.name
$version = $package.version
if (-not $publisher -or -not $name -or -not $version) {
    throw 'package.json missing publisher, name, or version.'
}

$extId = "$publisher.$name"
$extFolderName = "$extId-$version"

$useInsiders = $Insiders -or (-not $Stable)
$extensionsRoot = if (-not $useInsiders) {
    Join-Path $env:USERPROFILE '.vscode\extensions'
} else {
    Join-Path $env:USERPROFILE '.vscode-insiders\extensions'
}

if (-not (Test-Path -LiteralPath $extensionsRoot)) {
    New-Item -ItemType Directory -Path $extensionsRoot | Out-Null
}

Get-ChildItem -LiteralPath $extensionsRoot -Filter "$extId*" | ForEach-Object {
    if ($Force) {
        Remove-Item -LiteralPath $_.FullName -Recurse -Force
    } else {
        Write-Host "Existing extension found: $($_.FullName)" -ForegroundColor Yellow
        Write-Host "Re-run with -Force to replace it." -ForegroundColor Yellow
        exit 1
    }
}

$target = Join-Path $extensionsRoot $extFolderName

Write-Host "Linking $target -> $repoRoot" -ForegroundColor Cyan
if ($IsWindows) {
    New-Item -ItemType Junction -Path $target -Target $repoRoot | Out-Null
} else {
    New-Item -ItemType SymbolicLink -Path $target -Target $repoRoot | Out-Null
}

Write-Host "Installed dev link for $extId" -ForegroundColor Green
Write-Host "Run 'npm run compile' after changes, then reload the VS Code window." -ForegroundColor Green
