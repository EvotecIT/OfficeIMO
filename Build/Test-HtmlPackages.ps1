param(
    [string] $Version = '3.3.0'
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$workingPath = Join-Path ([System.IO.Path]::GetTempPath()) ('officeimo-html-package-smoke-' + [Guid]::NewGuid().ToString('N'))
$feedPath = Join-Path $workingPath 'feed'
$configPath = Join-Path $workingPath 'nuget\nuget.config'
$packagesPath = Join-Path $workingPath 'packages'
New-Item -ItemType Directory -Path $feedPath -Force | Out-Null

try {
    $projects = @(
        'OfficeIMO.Core/OfficeIMO.Core.csproj',
        'OfficeIMO.IWork/OfficeIMO.IWork.csproj',
        'OfficeIMO.Html/OfficeIMO.Html.csproj',
        'OfficeIMO.Word/OfficeIMO.Word.csproj',
        'OfficeIMO.Word.Html/OfficeIMO.Word.Html.csproj',
        'OfficeIMO.Excel/OfficeIMO.Excel.csproj',
        'OfficeIMO.Excel.Html/OfficeIMO.Excel.Html.csproj',
        'OfficeIMO.PowerPoint/OfficeIMO.PowerPoint.csproj',
        'OfficeIMO.PowerPoint.Html/OfficeIMO.PowerPoint.Html.csproj',
        'OfficeIMO.Rtf/OfficeIMO.Rtf.csproj',
        'OfficeIMO.Html.Rtf/OfficeIMO.Html.Rtf.csproj',
        'OfficeIMO.Pdf/OfficeIMO.Pdf.csproj',
        'OfficeIMO.Html.Pdf/OfficeIMO.Html.Pdf.csproj'
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
        '--property:EnableOfficeIMOHtmlPackageSmoke=true',
        "--property:OfficeIMOHtmlPackageVersion=$Version"
    )
    $projectPath = 'Build/PackageSmoke/OfficeIMO.Html/OfficeIMO.Html.PackageSmoke.csproj'
    dotnet restore $projectPath @properties --configfile $configPath --packages $packagesPath --no-http-cache --force-evaluate
    if ($LASTEXITCODE -ne 0) { throw 'Packed HTML consumer restore failed.' }

    $frameworks = if ($IsWindows) { @('net472', 'net8.0', 'net10.0') } else { @('net8.0', 'net10.0') }
    foreach ($framework in $frameworks) {
        dotnet run --project $projectPath --configuration Release --framework $framework --no-restore @properties
        if ($LASTEXITCODE -ne 0) { throw "Packed HTML consumer failed on $framework." }
    }
} finally {
    if (Test-Path -LiteralPath $workingPath) {
        Remove-Item -LiteralPath $workingPath -Recurse -Force
    }
}
