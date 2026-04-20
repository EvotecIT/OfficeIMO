[CmdletBinding()]
param(
    [string]$Configuration = "Release",
    [switch]$SkipTests
)

$ErrorActionPreference = 'Stop'

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -Path $scriptRoot

function Require-Command {
    param([string]$Name)
    if (-not (Get-Command $Name -ErrorAction SilentlyContinue)) {
        throw "Required command '$Name' not found."
    }
}

Require-Command -Name 'node'
Require-Command -Name 'npm'
Require-Command -Name 'dotnet'

Write-Host "OfficeIMO Markup VS Code build - $Configuration" -ForegroundColor Green

dotnet build ..\OfficeIMO.Markup.Cli\OfficeIMO.Markup.Cli.csproj -c $Configuration --framework net8.0 --no-restore -m:1 -p:BuildInParallel=false -p:UseSharedCompilation=false --nologo --verbosity minimal

try {
    npm ci
} catch {
    Write-Warning "npm ci failed. Retrying with npm install..."
    npm install
}

npm run compile

if (-not $SkipTests) {
    Write-Host "No extension test suite is wired yet." -ForegroundColor Yellow
}
