param(
    [string] $Version = '3.6.0-smoke'
)

# Packs OfficeIMO.Pdf.Testing and drives the same invoice snapshot through xUnit, NUnit and MSTest consumers:
# the first run fails with received files, approval makes it pass, and a changed table cell fails with a
# structure diff and an HTML review gallery.
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
# Consumers are expected to fail before approval, so a non-zero dotnet test exit is checked, not thrown.
$PSNativeCommandUseErrorActionPreference = $false
$repositoryRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
$workingPath = Join-Path ([System.IO.Path]::GetTempPath()) ('officeimo-pdf-testing-package-smoke-' + [Guid]::NewGuid().ToString('N'))
$feedPath = Join-Path $workingPath 'feed'
$buildPath = Join-Path $workingPath 'build'
$configPath = Join-Path $workingPath 'nuget/nuget.config'
$packagesPath = Join-Path $workingPath 'packages'
$smokeRoot = Join-Path $repositoryRoot 'Build/PackageSmoke/OfficeIMO.Pdf.Testing'
$consumers = @('Xunit', 'NUnit', 'MSTest')
New-Item -ItemType Directory -Path $feedPath -Force | Out-Null

function Invoke-SmokeTest {
    param([string] $Project, [string] $Framework, [bool] $ExpectFailure)
    $output = & dotnet test $Project --configuration Release --framework $Framework --no-build --logger 'console;verbosity=normal' 2>&1 | Out-String
    $failed = $LASTEXITCODE -ne 0
    if ($failed -ne $ExpectFailure) {
        Write-Host $output
        throw "$Project on $Framework $(if ($ExpectFailure) { 'passed but should have failed' } else { 'failed' })."
    }
    return $output
}

function Assert-Contains {
    param([string] $Text, [string] $Expected, [string] $Context)
    if (-not $Text.Contains($Expected)) {
        Write-Host $Text
        throw "$Context did not contain '$Expected'."
    }
}

Push-Location $repositoryRoot
try {
    # Package the complete public framework graph on every host. Omitting net472
    # makes API compatibility compare its baseline against netstandard2.0 instead.
    $packFrameworks = '--property:TargetFrameworks="netstandard2.0;net8.0;net10.0;net472"'
    foreach ($project in @('OfficeIMO.Core/OfficeIMO.Core.csproj', 'OfficeIMO.Pdf/OfficeIMO.Pdf.csproj', 'OfficeIMO.Pdf.Testing/OfficeIMO.Pdf.Testing.csproj')) {
        dotnet restore $project $packFrameworks --artifacts-path $buildPath --no-http-cache
        if ($LASTEXITCODE -ne 0) { throw "Restore failed for $project." }
        dotnet pack $project --configuration Release --artifacts-path $buildPath --no-restore --output $feedPath --property:PackageVersion=$Version $packFrameworks
        if ($LASTEXITCODE -ne 0) { throw "Pack failed for $project." }
    }

    $configDirectory = Split-Path -Parent $configPath
    dotnet new nugetconfig --output $configDirectory --force
    if ($LASTEXITCODE -ne 0) { throw 'NuGet configuration creation failed.' }
    dotnet nuget add source $feedPath --name OfficeIMOLocal --configfile $configPath
    if ($LASTEXITCODE -ne 0) { throw 'Local package source registration failed.' }

    $properties = @(
        '--property:EnableOfficeIMOPdfTestingPackageSmoke=true',
        "--property:OfficeIMOPdfTestingPackageVersion=$Version"
    )
    $frameworks = if ($IsWindows) { @('net8.0', 'net10.0', 'net472') } else { @('net8.0', 'net10.0') }

    foreach ($consumer in $consumers) {
        $consumerPath = Join-Path $smokeRoot $consumer
        $project = Join-Path $consumerPath "OfficeIMO.Pdf.Testing.PackageSmoke.$consumer.csproj"
        $snapshots = Join-Path $consumerPath '__pdfsnapshots__'
        if (Test-Path -LiteralPath $snapshots) { Remove-Item -LiteralPath $snapshots -Recurse -Force }

        dotnet restore $project @properties --configfile $configPath --packages $packagesPath --no-http-cache --force-evaluate
        if ($LASTEXITCODE -ne 0) { throw "Packed $consumer consumer restore failed." }
        dotnet build $project --configuration Release --no-restore @properties
        if ($LASTEXITCODE -ne 0) { throw "Packed $consumer consumer failed to build." }

        $first = $frameworks[0]
        $output = Invoke-SmokeTest $project $first $true
        Assert-Contains $output 'has no approved baseline' "$consumer first run"
        foreach ($file in @('invoice.received.structure.json', 'invoice.received.page-001.png')) {
            if (-not (Test-Path -LiteralPath (Join-Path $snapshots $file))) { throw "$consumer first run did not write $file." }
        }

        $env:OFFICEIMO_PDF_SNAPSHOT_ACCEPT = '1'
        try { Invoke-SmokeTest $project $first $false | Out-Null } finally { Remove-Item Env:OFFICEIMO_PDF_SNAPSHOT_ACCEPT }
        if (Get-ChildItem -LiteralPath $snapshots -Filter '*.received.*') { throw "$consumer approval left received files." }
        foreach ($framework in $frameworks) { Invoke-SmokeTest $project $framework $false | Out-Null }

        $env:PDF_SNAPSHOT_SMOKE_TOTAL = '1,256.00'
        try { $output = Invoke-SmokeTest $project $first $true } finally { Remove-Item Env:PDF_SNAPSHOT_SMOKE_TOTAL }
        if ($output -notmatch '(?m)^\s*-\s+"\| Total \| 1,250\.00 \|"' -or $output -notmatch '(?m)^\s*\+\s+"\| Total \| 1,256\.00 \|"') {
            Write-Host $output
            throw "$consumer changed-cell failure did not show the table cell diff."
        }
        Assert-Contains $output 'invoice.received.html' "$consumer changed-cell message"
        $gallery = Join-Path $snapshots 'invoice.received.html'
        if (-not (Test-Path -LiteralPath $gallery)) { throw "$consumer did not write the review gallery." }
        Assert-Contains (Get-Content -LiteralPath $gallery -Raw) 'Differs' "$consumer review gallery"
        Write-Host "$consumer consumer: missing baseline failed, approval passed on $($frameworks -join ', '), changed cell reported."
    }
} finally {
    Pop-Location
    foreach ($consumer in $consumers) {
        $snapshots = Join-Path (Join-Path $smokeRoot $consumer) '__pdfsnapshots__'
        if (Test-Path -LiteralPath $snapshots) { Remove-Item -LiteralPath $snapshots -Recurse -Force }
    }
    if (Test-Path -LiteralPath $workingPath) {
        Remove-Item -LiteralPath $workingPath -Recurse -Force
    }
}
