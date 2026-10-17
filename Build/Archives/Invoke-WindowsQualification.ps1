[CmdletBinding()]
param(
    [Parameter(Mandatory)][string] $OutputPath,
    [Parameter(Mandatory)][string] $FeedPath,
    [Parameter(Mandatory)][string] $Version
)

$ErrorActionPreference = 'Stop'
if (-not $IsWindows) { throw 'This qualification requires Windows.' }
$repo = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$checkoutCommit = & git -C $repo rev-parse --verify HEAD
if ($LASTEXITCODE -ne 0 -or $checkoutCommit -notmatch '^[0-9a-f]{40}$') { throw 'Cannot identify the qualification checkout commit.' }
$OutputPath = [IO.Path]::GetFullPath($OutputPath)
if (Test-Path -LiteralPath $OutputPath) { throw 'Use a fresh task-owned output directory.' }
$reports = Join-Path $OutputPath 'reports'
$null = New-Item -ItemType Directory -Path $reports
$env:TEMP = Join-Path $OutputPath 'tmp'; $env:TMP = $env:TEMP
$null = New-Item -ItemType Directory -Path $env:TEMP
$project = Join-Path $PSScriptRoot 'PackageConsumer/PackageConsumer.csproj'
$corpus = Join-Path $repo 'OfficeIMO.Archives.Tests/Fixtures/IndependentZip'
$msi = Join-Path $OutputPath '7z2604-x64.msi'
Invoke-WebRequest -Uri 'https://github.com/ip7z/7zip/releases/download/26.04/7z2604-x64.msi' -OutFile $msi
if ((Get-FileHash $msi -Algorithm SHA256).Hash -ne '0b01334a418654293513449f61d6bfc99e5196ca371e3c3dc961eda57cd535c6') { throw 'Pinned MSI hash mismatch.' }
$cache = Join-Path $OutputPath 'private-package-cache'
$artifacts = Join-Path $OutputPath 'consumer-artifacts'
$configDirectory = Join-Path $OutputPath 'nuget'
dotnet new nugetconfig --output $configDirectory
if ($LASTEXITCODE -ne 0) { throw 'Private NuGet configuration creation failed.' }
$configPath = Join-Path $configDirectory 'nuget.config'
dotnet nuget add source $FeedPath --name OfficeIMOQualification --configfile $configPath
if ($LASTEXITCODE -ne 0) { throw 'Private package source configuration failed.' }
dotnet restore $project "-p:OfficeIMOQualificationVersion=$Version" --configfile $configPath --packages $cache --artifacts-path $artifacts
if ($LASTEXITCODE -ne 0) { throw 'Private package restore failed.' }
dotnet build $project -c Release --no-restore "-p:OfficeIMOQualificationVersion=$Version" "-p:RestorePackagesPath=$cache" --artifacts-path $artifacts -m:1
if ($LASTEXITCODE -ne 0) { throw 'Packed consumer build failed.' }
foreach ($framework in @('net8.0', 'net10.0', 'net472')) {
    $arguments = @($corpus, (Join-Path $OutputPath "extracted-$framework"), (Join-Path $reports "$framework.json"), $msi,
        (Join-Path $PSScriptRoot 'Evidence/7zip-26.04-msi-payloads.json'),
        (Join-Path $repo 'OfficeIMO.OneNote.Tests/Fixtures/makecab-lzx-testOneNote2016.cab'),
        (Join-Path $repo 'OfficeIMO.OneNote.Tests/Fixtures/testOneNote2016.one'))
    $binary = Join-Path $artifacts "bin/PackageConsumer/release_$framework/PackageConsumer"
    if ($framework -eq 'net472') { & "$binary.exe" @arguments }
    else { & dotnet "$binary.dll" @arguments }
    if ($LASTEXITCODE -ne 0) { throw "Packed consumer failed: $framework" }
}
@{ checkoutCommit = $checkoutCommit; sourceHead = $env:OFFICEIMO_QUALIFICATION_HEAD
    runUrl = "https://github.com/$env:GITHUB_REPOSITORY/actions/runs/$env:GITHUB_RUN_ID"
    packageVersion = $Version; os = [Environment]::OSVersion.VersionString
    packages = @(@('OfficeIMO.Core', 'OfficeIMO.Archives', 'OfficeIMO.Reader.Core', 'OfficeIMO.Reader.Zip') | ForEach-Object {
        $package = Join-Path $FeedPath "$_.$Version.nupkg"
        @{ file = [IO.Path]::GetFileName($package); bytes = (Get-Item $package).Length
            sha256 = (Get-FileHash $package -Algorithm SHA256).Hash.ToLowerInvariant() }
    })
    boundary = 'Unsigned private package runtime proof; Windows Shell ZIP producer, not interactive Explorer UI automation. No installer or extracted executable run.'
} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $reports 'context.json') -Encoding utf8

# Native Shell production is separate: a headless producer failure must not prevent the runtime proof above.
$native = Join-Path $OutputPath 'native'
& pwsh -NoProfile -STA -File (Join-Path $PSScriptRoot 'New-WindowsShellZipFixture.ps1') -OutputPath $native
$nativeExitCode = $LASTEXITCODE
foreach ($name in @('producer-attempt.json', 'windows-shell-polish.zip')) {
    $path = Join-Path $native $name
    if (Test-Path -LiteralPath $path) { Copy-Item -LiteralPath $path -Destination $reports }
}
if ($nativeExitCode -ne 0) { throw 'Native Windows Shell producer failed; base packed runtime and producer diagnostics are retained.' }
$fixture = Get-Content -Raw -LiteralPath (Join-Path $native 'fixture.json') | ConvertFrom-Json
$nativeCorpus = Join-Path $OutputPath 'native-corpus'
$null = New-Item -ItemType Directory -Path $nativeCorpus
Copy-Item -LiteralPath (Join-Path $native $fixture.file) -Destination $nativeCorpus
Copy-Item -LiteralPath (Join-Path $native $fixture.file) -Destination $reports
Copy-Item -LiteralPath (Join-Path $native 'fixture.json') -Destination (Join-Path $reports 'windows-producer.json')
@{ fixtures = @($fixture) } | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $nativeCorpus 'manifest.json') -Encoding utf8
foreach ($framework in @('net8.0', 'net10.0', 'net472')) {
    $arguments = @($nativeCorpus, (Join-Path $OutputPath "native-extracted-$framework"), (Join-Path $reports "native-$framework.json"))
    $binary = Join-Path $artifacts "bin/PackageConsumer/release_$framework/PackageConsumer"
    if ($framework -eq 'net472') { & "$binary.exe" @arguments }
    else { & dotnet "$binary.dll" @arguments }
    if ($LASTEXITCODE -ne 0) { throw "Native Shell fixture consumer failed: $framework" }
}
