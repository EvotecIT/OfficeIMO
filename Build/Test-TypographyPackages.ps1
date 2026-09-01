param(
    [string] $Version = '3.3.0'
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$repositoryRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
$workingPath = Join-Path ([System.IO.Path]::GetTempPath()) ('officeimo-typography-package-smoke-' + [Guid]::NewGuid().ToString('N'))
$feedPath = Join-Path $workingPath 'feed'
$configPath = Join-Path $workingPath 'nuget\nuget.config'
$packagesPath = Join-Path $workingPath 'packages'
New-Item -ItemType Directory -Path $feedPath -Force | Out-Null
Push-Location $repositoryRoot
try {
    $projects = @(
        'OfficeIMO.Core/OfficeIMO.Core.csproj',
        'OfficeIMO.Drawing.HarfBuzz/OfficeIMO.Drawing.HarfBuzz.csproj'
    )

    foreach ($project in $projects) {
        dotnet restore $project --no-http-cache
        if ($LASTEXITCODE -ne 0) { throw "Restore failed for $project." }
        dotnet pack $project --configuration Release --no-restore --output $feedPath --property:PackageVersion=$Version
        if ($LASTEXITCODE -ne 0) { throw "Pack failed for $project." }
    }

    $configDirectory = Split-Path -Parent $configPath
    dotnet new nugetconfig --output $configDirectory --force
    if ($LASTEXITCODE -ne 0) { throw 'NuGet configuration creation failed.' }
    dotnet nuget add source $feedPath --name OfficeIMOLocal --configfile $configPath
    if ($LASTEXITCODE -ne 0) { throw 'Local package source registration failed.' }

    $properties = @(
        '--property:EnableOfficeIMOTypographyPackageSmoke=true',
        "--property:OfficeIMOTypographyPackageVersion=$Version"
    )
    $projectPath = 'Build/PackageSmoke/OfficeIMO.Typography/OfficeIMO.Typography.PackageSmoke.csproj'
    dotnet restore $projectPath @properties --configfile $configPath --packages $packagesPath --no-http-cache --force-evaluate
    if ($LASTEXITCODE -ne 0) { throw 'Packed typography consumer restore failed.' }

    dotnet build $projectPath --configuration Release --framework netstandard2.0 --no-restore @properties
    if ($LASTEXITCODE -ne 0) { throw 'Packed typography consumer failed to compile on netstandard2.0.' }

    $frameworks = if ($IsWindows) { @('net472', 'net8.0', 'net10.0') } else { @('net8.0', 'net10.0') }
    foreach ($framework in $frameworks) {
        dotnet run --project $projectPath --configuration Release --framework $framework --no-restore @properties
        if ($LASTEXITCODE -ne 0) { throw "Packed typography consumer failed on $framework." }
    }
} finally {
    Pop-Location
    if (Test-Path -LiteralPath $workingPath) {
        Remove-Item -LiteralPath $workingPath -Recurse -Force
    }
}
