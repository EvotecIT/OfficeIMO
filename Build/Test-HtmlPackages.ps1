param(
    [string] $Version = '3.3.0',
    [string] $ArtifactsPath
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$retainArtifacts = -not [string]::IsNullOrWhiteSpace($ArtifactsPath)
$workingPath = if ($retainArtifacts) { [System.IO.Path]::GetFullPath($ArtifactsPath) } else {
    Join-Path ([System.IO.Path]::GetTempPath()) ('officeimo-html-package-smoke-' + [Guid]::NewGuid().ToString('N'))
}
if (Test-Path -LiteralPath $workingPath) { throw "Choose a new artifact directory: $workingPath" }
$feedPath = Join-Path $workingPath 'feed'
$configPath = Join-Path $workingPath 'nuget\nuget.config'
$packagesPath = Join-Path $workingPath 'packages'
New-Item -ItemType Directory -Path $feedPath -Force | Out-Null

try {
    $projects = @(
        'OfficeIMO.Core/OfficeIMO.Core.csproj',
        'OfficeIMO.Html.Core/OfficeIMO.Html.Core.csproj',
        'OfficeIMO.Html.AngleSharp/OfficeIMO.Html.AngleSharp.csproj',
        'OfficeIMO.Html/OfficeIMO.Html.csproj',
        'OfficeIMO.Markdown/OfficeIMO.Markdown.csproj',
        'OfficeIMO.Markdown.Html/OfficeIMO.Markdown.Html.csproj',
        'OfficeIMO.Word/OfficeIMO.Word.csproj',
        'OfficeIMO.Word.Html/OfficeIMO.Word.Html.csproj',
        'OfficeIMO.Excel/OfficeIMO.Excel.csproj',
        'OfficeIMO.Excel.Html/OfficeIMO.Excel.Html.csproj',
        'OfficeIMO.PowerPoint/OfficeIMO.PowerPoint.csproj',
        'OfficeIMO.PowerPoint.Html/OfficeIMO.PowerPoint.Html.csproj',
        'OfficeIMO.Rtf/OfficeIMO.Rtf.csproj',
        'OfficeIMO.Html.Rtf/OfficeIMO.Html.Rtf.csproj',
        'OfficeIMO.Pdf/OfficeIMO.Pdf.csproj',
        'OfficeIMO.Html.Pdf/OfficeIMO.Html.Pdf.csproj',
        'OfficeIMO.Email/OfficeIMO.Email.csproj',
        'OfficeIMO.Mhtml/OfficeIMO.Mhtml.csproj',
        'OfficeIMO.Mhtml.Pdf/OfficeIMO.Mhtml.Pdf.csproj'
    )
    # Package the complete public framework graph on every host. Omitting net472
    # makes API compatibility compare its baseline against netstandard2.0 instead.
    $packFrameworks = '--property:TargetFrameworks="netstandard2.0;net8.0;net10.0;net472"'
    foreach ($project in $projects) {
        dotnet restore $project $packFrameworks --no-http-cache
        if ($LASTEXITCODE -ne 0) { throw "Restore failed for $project." }
        dotnet pack $project --configuration Release --no-restore --output $feedPath --property:PackageVersion=$Version $packFrameworks
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
    $consumerProjects = @(
        'Build/PackageSmoke/OfficeIMO.Html.Document/OfficeIMO.Html.Document.PackageSmoke.csproj',
        'Build/PackageSmoke/OfficeIMO.Html/OfficeIMO.Html.PackageSmoke.csproj'
    )
    $frameworks = if ($IsWindows) { @('net472', 'net8.0', 'net10.0') } else { @('net8.0', 'net10.0') }
    foreach ($projectPath in $consumerProjects) {
        dotnet restore $projectPath @properties --configfile $configPath --packages $packagesPath --no-http-cache --force-evaluate
        if ($LASTEXITCODE -ne 0) { throw "Packed HTML consumer restore failed for $projectPath." }
        foreach ($framework in $frameworks) {
            dotnet run --project $projectPath --configuration Release --framework $framework --no-restore @properties
            if ($LASTEXITCODE -ne 0) { throw "Packed HTML consumer failed for $projectPath on $framework." }
        }
    }
} finally {
    if (-not $retainArtifacts -and (Test-Path -LiteralPath $workingPath)) {
        # NuGet creates dot-prefixed metadata that PowerShell treats as hidden on Unix.
        # This path is the fresh temporary tree created by this invocation.
        [System.IO.Directory]::Delete($workingPath, $true)
    }
}
