param(
    [string] $Version = '3.2.5-typography-local',
    [switch] $RequireSixLabors
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$repositoryRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
$workingPath = Join-Path ([System.IO.Path]::GetTempPath()) ('officeimo-typography-package-smoke-' + [Guid]::NewGuid().ToString('N'))
$feedPath = Join-Path $workingPath 'feed'
$configPath = Join-Path $workingPath 'nuget\nuget.config'
$packagesPath = Join-Path $workingPath 'packages'
$hasSixLaborsLicense = -not [string]::IsNullOrWhiteSpace($env:SIXLABORS_LICENSE_KEY)
$previousMsBuildLicenseKey = $env:SixLaborsLicenseKey

if ($RequireSixLabors -and -not $hasSixLaborsLicense) {
    throw 'A permanent Six Labors license is required for this trusted package gate. Configure the SIXLABORS_LICENSE_KEY secret with the complete supplied license value.'
}

New-Item -ItemType Directory -Path $feedPath -Force | Out-Null
Push-Location $repositoryRoot
try {
    if ($hasSixLaborsLicense) {
        $env:SixLaborsLicenseKey = $env:SIXLABORS_LICENSE_KEY
    }

    $projects = @(
        'OfficeIMO.Core/OfficeIMO.Core.csproj',
        'OfficeIMO.Drawing.HarfBuzz/OfficeIMO.Drawing.HarfBuzz.csproj'
    )
    if ($hasSixLaborsLicense) {
        $projects += 'OfficeIMO.Drawing.SixLabors/OfficeIMO.Drawing.SixLabors.csproj'
    }

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
        "--property:OfficeIMOTypographyPackageVersion=$Version",
        "--property:IncludeOfficeIMOSixLaborsPackage=$($hasSixLaborsLicense.ToString().ToLowerInvariant())"
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

    if (-not $hasSixLaborsLicense) {
        Write-Warning 'SixLabors package smoke was skipped because SIXLABORS_LICENSE_KEY is not configured.'
    }
} finally {
    $env:SixLaborsLicenseKey = $previousMsBuildLicenseKey
    Pop-Location
    if (Test-Path -LiteralPath $workingPath) {
        Remove-Item -LiteralPath $workingPath -Recurse -Force
    }
}
