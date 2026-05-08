[CmdletBinding()]
param(
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Release',

    [string]$Framework = 'net8.0',

    [string]$OutputDirectory = 'dist',

    [switch]$SkipNpmCi,

    [switch]$SkipRestore,

    [switch]$PublishMarketplace,

    [switch]$PreRelease,

    [string]$VscePat = $env:VSCE_PAT
)

$ErrorActionPreference = 'Stop'

function Assert-ChildPath {
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [string]$Parent
    )

    $resolved = [System.IO.Path]::GetFullPath($Path)
    $resolvedParent = [System.IO.Path]::GetFullPath($Parent)
    $parentWithSeparator = $resolvedParent.TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar

    if ($resolved -ne $resolvedParent -and -not $resolved.StartsWith($parentWithSeparator, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to operate on '$resolved' because it is outside '$resolvedParent'."
    }

    return $resolved
}

function Invoke-Tool {
    param(
        [Parameter(Mandatory)]
        [string]$FilePath,

        [Parameter()]
        [string[]]$ArgumentList = @()
    )

    & $FilePath @ArgumentList
    if ($LASTEXITCODE -ne 0) {
        throw "'$FilePath $($ArgumentList -join ' ')' failed with exit code $LASTEXITCODE."
    }
}

function Get-VsceCommand {
    param(
        [Parameter(Mandatory)]
        [string]$ExtensionRoot
    )

    $windowsVsce = Join-Path $ExtensionRoot 'node_modules/.bin/vsce.cmd'
    if (Test-Path -LiteralPath $windowsVsce) {
        return @{
            FilePath = $windowsVsce
            Prefix = @()
        }
    }

    $posixVsce = Join-Path $ExtensionRoot 'node_modules/.bin/vsce'
    if (Test-Path -LiteralPath $posixVsce) {
        return @{
            FilePath = $posixVsce
            Prefix = @()
        }
    }

    $vsceMain = Join-Path $ExtensionRoot 'node_modules/@vscode/vsce/out/main.js'
    if (Test-Path -LiteralPath $vsceMain) {
        $node = (Get-Command node -ErrorAction Stop).Source
        return @{
            FilePath = $node
            Prefix = @($vsceMain)
        }
    }

    throw 'VSCE was not found in node_modules. Run npm ci first.'
}

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$extensionRoot = (Resolve-Path -LiteralPath (Join-Path $scriptRoot '..')).Path
$repoRoot = (Resolve-Path -LiteralPath (Join-Path $extensionRoot '..')).Path
$packagePath = Join-Path $extensionRoot 'package.json'
$cliProject = Join-Path $repoRoot 'OfficeIMO.Markup.Cli/OfficeIMO.Markup.Cli.csproj'

if (-not (Test-Path -LiteralPath $packagePath)) {
    throw "package.json not found at $packagePath."
}

if (-not (Test-Path -LiteralPath $cliProject)) {
    throw "OfficeIMO.Markup.Cli project not found at $cliProject."
}

Get-Command node -ErrorAction Stop | Out-Null
Get-Command npm -ErrorAction Stop | Out-Null
Get-Command dotnet -ErrorAction Stop | Out-Null

Push-Location $extensionRoot
try {
    if (-not $SkipNpmCi) {
        if (Test-Path -LiteralPath (Join-Path $extensionRoot 'package-lock.json')) {
            Write-Host 'Installing extension dependencies with npm ci...' -ForegroundColor Cyan
            Invoke-Tool -FilePath 'npm' -ArgumentList @('ci')
        } else {
            Write-Host 'Installing extension dependencies with npm install...' -ForegroundColor Cyan
            Invoke-Tool -FilePath 'npm' -ArgumentList @('install')
        }
    }

    $publishRoot = Assert-ChildPath -Path (Join-Path $extensionRoot '.tmp/cli-publish') -Parent $extensionRoot
    if (Test-Path -LiteralPath $publishRoot) {
        Remove-Item -LiteralPath $publishRoot -Recurse -Force
    }
    New-Item -ItemType Directory -Path $publishRoot | Out-Null

    $dotnetArgs = @(
        'publish',
        $cliProject,
        '-c', $Configuration,
        '-f', $Framework,
        '-o', $publishRoot,
        '--nologo',
        '--verbosity', 'minimal',
        '-m:1',
        '-nr:false',
        '-p:BuildInParallel=false',
        '-p:UseSharedCompilation=false',
        '-p:DebugType=embedded'
    )
    if ($SkipRestore) {
        $dotnetArgs += '--no-restore'
    }

    Write-Host "Publishing OfficeIMO.Markup.Cli ($Configuration, $Framework)..." -ForegroundColor Cyan
    Invoke-Tool -FilePath 'dotnet' -ArgumentList $dotnetArgs

    $bundledCli = Assert-ChildPath -Path (Join-Path $extensionRoot 'tools/OfficeIMO.Markup.Cli') -Parent $extensionRoot
    if (Test-Path -LiteralPath $bundledCli) {
        Remove-Item -LiteralPath $bundledCli -Recurse -Force
    }
    New-Item -ItemType Directory -Path $bundledCli | Out-Null
    Copy-Item -Path (Join-Path $publishRoot '*') -Destination $bundledCli -Recurse -Force
    Remove-Item -LiteralPath $publishRoot -Recurse -Force

    Write-Host 'Compiling VS Code extension...' -ForegroundColor Cyan
    Invoke-Tool -FilePath 'npm' -ArgumentList @('run', 'compile')

    $package = Get-Content -LiteralPath $packagePath -Raw | ConvertFrom-Json
    $outputRoot = if ([System.IO.Path]::IsPathRooted($OutputDirectory)) {
        $OutputDirectory
    } else {
        Join-Path $extensionRoot $OutputDirectory
    }
    $outputRoot = Assert-ChildPath -Path $outputRoot -Parent $extensionRoot
    New-Item -ItemType Directory -Path $outputRoot -Force | Out-Null

    $vsixPath = Join-Path $outputRoot ("{0}-{1}.vsix" -f $package.name, $package.version)
    if (Test-Path -LiteralPath $vsixPath) {
        Remove-Item -LiteralPath $vsixPath -Force
    }

    $vsce = Get-VsceCommand -ExtensionRoot $extensionRoot
    $packageArgs = @($vsce.Prefix + @('package', '--allow-missing-repository', '--out', $vsixPath))
    if ($PreRelease) {
        $packageArgs += '--pre-release'
    }

    Write-Host 'Packaging VSIX...' -ForegroundColor Cyan
    Invoke-Tool -FilePath $vsce.FilePath -ArgumentList $packageArgs

    if ($PublishMarketplace) {
        if ([string]::IsNullOrWhiteSpace($VscePat)) {
            throw 'VSCE_PAT is required when publishing to the Visual Studio Marketplace.'
        }

        $publishArgs = @($vsce.Prefix + @('publish', '--packagePath', $vsixPath))
        if ($PreRelease) {
            $publishArgs += '--pre-release'
        }

        Write-Host 'Publishing VSIX to the Visual Studio Marketplace...' -ForegroundColor Cyan
        $previousVscePat = $env:VSCE_PAT
        try {
            $env:VSCE_PAT = $VscePat
            Invoke-Tool -FilePath $vsce.FilePath -ArgumentList $publishArgs
        } finally {
            $env:VSCE_PAT = $previousVscePat
        }
    }

    Write-Host "VSIX: $vsixPath" -ForegroundColor Green
} finally {
    Pop-Location
}
