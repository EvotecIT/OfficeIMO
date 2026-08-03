param(
    [string] $Version = '3.1.0'
)

$ErrorActionPreference = 'Stop'
$PSNativeCommandUseErrorActionPreference = $true
if ($Version -notmatch '^\d+\.\d+\.\d+$') { throw 'Version must be a public three-part version.' }

$repositoryRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$temporaryRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
$workingPath = [IO.Path]::GetFullPath((Join-Path $temporaryRoot ('officeimo-google-package-smoke-' + [Guid]::NewGuid().ToString('N'))))
if (-not $workingPath.StartsWith($temporaryRoot, [StringComparison]::OrdinalIgnoreCase)) { throw 'Package-smoke path escaped the temporary directory.' }
$feed = Join-Path $workingPath 'feed'
$packages = Join-Path $workingPath 'packages'
$configPath = Join-Path $workingPath 'NuGet.Config'

try {
    New-Item -ItemType Directory -Path $feed | Out-Null
    $projects = @(
        'OfficeIMO.GoogleWorkspace/OfficeIMO.GoogleWorkspace.csproj',
        'OfficeIMO.GoogleWorkspace.Drive/OfficeIMO.GoogleWorkspace.Drive.csproj',
        'OfficeIMO.GoogleWorkspace.Sync/OfficeIMO.GoogleWorkspace.Sync.csproj',
        'OfficeIMO.GoogleWorkspace.Auth.GoogleApis/OfficeIMO.GoogleWorkspace.Auth.GoogleApis.csproj'
    )
    foreach ($project in $projects) {
        dotnet pack (Join-Path $repositoryRoot $project) -c Release -p:PackageVersion=$Version -o $feed
    }

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    foreach ($id in @('OfficeIMO.GoogleWorkspace', 'OfficeIMO.GoogleWorkspace.Drive', 'OfficeIMO.GoogleWorkspace.Sync')) {
        $package = Get-Item (Join-Path $feed "$id.$Version.nupkg")
        $archive = [IO.Compression.ZipFile]::OpenRead($package.FullName)
        try {
            $entries = @($archive.Entries | Where-Object { $_.FullName.EndsWith('.nuspec', [StringComparison]::OrdinalIgnoreCase) })
            if ($entries.Count -ne 1) { throw "$id package must contain exactly one nuspec." }
            $entry = $entries[0]
            $reader = [IO.StreamReader]::new($entry.Open())
            try { $nuspec = $reader.ReadToEnd() } finally { $reader.Dispose() }
            if ($nuspec -match 'Google\.Apis') { throw "$id unexpectedly depends on a Google client SDK package." }
        } finally { $archive.Dispose() }
    }
    $adapterPackage = Get-Item (Join-Path $feed "OfficeIMO.GoogleWorkspace.Auth.GoogleApis.$Version.nupkg")
    $adapterArchive = [IO.Compression.ZipFile]::OpenRead($adapterPackage.FullName)
    try {
        $adapterEntries = @($adapterArchive.Entries | Where-Object { $_.FullName.EndsWith('.nuspec', [StringComparison]::OrdinalIgnoreCase) })
        if ($adapterEntries.Count -ne 1) { throw 'The optional auth adapter package must contain exactly one nuspec.' }
        $adapterEntry = $adapterEntries[0]
        $adapterReader = [IO.StreamReader]::new($adapterEntry.Open())
        try { $adapterNuspec = $adapterReader.ReadToEnd() } finally { $adapterReader.Dispose() }
        if ($adapterNuspec -notmatch 'Google\.Apis\.Auth') { throw 'The optional auth adapter no longer declares Google.Apis.Auth.' }
    } finally { $adapterArchive.Dispose() }

    @"
<?xml version="1.0" encoding="utf-8"?>
<configuration><packageSources><clear /><add key="OfficeIMOLocal" value="$feed" /><add key="nuget.org" value="https://api.nuget.org/v3/index.json" /></packageSources><packageSourceMapping><packageSource key="OfficeIMOLocal"><package pattern="OfficeIMO.GoogleWorkspace*" /></packageSource><packageSource key="nuget.org"><package pattern="*" /></packageSource></packageSourceMapping></configuration>
"@ | Set-Content -LiteralPath $configPath -Encoding utf8

    $project = Join-Path $repositoryRoot 'Build/PackageSmoke/OfficeIMO.GoogleWorkspace/OfficeIMO.GoogleWorkspace.PackageSmoke.csproj'
    $properties = @('-p:EnableOfficeIMOGooglePackageSmoke=true', "-p:OfficeIMOGooglePackageVersion=$Version")
    dotnet restore $project @properties --configfile $configPath --packages $packages --no-cache --force-evaluate
    foreach ($framework in @('net472', 'net8.0', 'net10.0')) {
        dotnet run --project $project -c Release -f $framework --no-restore @properties
    }
} finally {
    if (Test-Path -LiteralPath $workingPath) { Remove-Item -LiteralPath $workingPath -Recurse -Force }
}
