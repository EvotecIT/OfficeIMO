param(
    [string] $HtmlTinkerXPackagePath,
    [string] $OfficeIMOVersion = '3.2.5-browser-local',
    [version] $MinimumHtmlTinkerXVersion = '3.0.1'
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

function Invoke-DotNet {
    param([Parameter(ValueFromRemainingArguments)][string[]] $Arguments)

    & dotnet @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "dotnet $($Arguments -join ' ') failed with exit code $LASTEXITCODE."
    }
}

function Get-NuGetIdentity {
    param([Parameter(Mandatory)][string] $PackagePath)

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [System.IO.Compression.ZipFile]::OpenRead($PackagePath)
    try {
        $nuspecs = @($archive.Entries | Where-Object {
                $_.FullName.EndsWith('.nuspec', [StringComparison]::OrdinalIgnoreCase)
            })
        if ($nuspecs.Count -ne 1) {
            throw "Package '$PackagePath' must contain exactly one nuspec."
        }

        $reader = [System.IO.StreamReader]::new($nuspecs[0].Open())
        try {
            [xml] $nuspec = $reader.ReadToEnd()
        } finally {
            $reader.Dispose()
        }

        $namespace = [System.Xml.XmlNamespaceManager]::new($nuspec.NameTable)
        $namespace.AddNamespace('n', $nuspec.DocumentElement.NamespaceURI)
        $metadata = $nuspec.SelectSingleNode('/n:package/n:metadata', $namespace)
        [pscustomobject]@{
            Id = [string] $metadata.id
            Version = [version] ([string] $metadata.version)
        }
    } finally {
        $archive.Dispose()
    }
}

$repositoryRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
$browserProjectPath = Join-Path $repositoryRoot 'OfficeIMO.Html.Pdf.Browser/OfficeIMO.Html.Pdf.Browser.csproj'
[xml] $browserProject = Get-Content -LiteralPath $browserProjectPath -Raw
$configuredVersionText = [string] $browserProject.Project.PropertyGroup.HtmlTinkerXVersion
if ([string]::IsNullOrWhiteSpace($configuredVersionText)) {
    throw 'OfficeIMO.Html.Pdf.Browser does not declare HtmlTinkerXVersion.'
}

$htmlTinkerXVersion = [version] $configuredVersionText
$resolvedHtmlTinkerXPackage = $null
if (-not [string]::IsNullOrWhiteSpace($HtmlTinkerXPackagePath)) {
    $resolvedHtmlTinkerXPackage = (Resolve-Path -LiteralPath $HtmlTinkerXPackagePath).Path
    $identity = Get-NuGetIdentity -PackagePath $resolvedHtmlTinkerXPackage
    if (-not $identity.Id.Equals('HtmlTinkerX', [StringComparison]::OrdinalIgnoreCase)) {
        throw "Package '$resolvedHtmlTinkerXPackage' is '$($identity.Id)', not HtmlTinkerX."
    }
    $htmlTinkerXVersion = $identity.Version
}

if ($htmlTinkerXVersion -lt $MinimumHtmlTinkerXVersion) {
    throw "OfficeIMO browser package proof requires HtmlTinkerX $MinimumHtmlTinkerXVersion or newer; selected version is $htmlTinkerXVersion."
}

$workingPath = Join-Path ([System.IO.Path]::GetTempPath()) (
    'oipb-' + [Guid]::NewGuid().ToString('N').Substring(0, 8))
$feedPath = Join-Path $workingPath 'f'
$configPath = Join-Path $workingPath 'n/nuget.config'
$packagesPath = Join-Path $workingPath 'p'
$artifactsPath = Join-Path $workingPath 'a'
New-Item -ItemType Directory -Path $feedPath -Force | Out-Null

Push-Location $repositoryRoot
try {
    if ($resolvedHtmlTinkerXPackage) {
        Copy-Item -LiteralPath $resolvedHtmlTinkerXPackage -Destination $feedPath
    }

    $configDirectory = Split-Path -Parent $configPath
    Invoke-DotNet new nugetconfig --output $configDirectory --force
    Invoke-DotNet nuget add source $feedPath --name OfficeIMOLocal --configfile $configPath

    foreach ($project in @(
            'OfficeIMO.Core/OfficeIMO.Core.csproj',
            'OfficeIMO.Pdf/OfficeIMO.Pdf.csproj')) {
        Invoke-DotNet restore $project --artifacts-path $artifactsPath --configfile $configPath --packages $packagesPath --no-http-cache --force-evaluate
        Invoke-DotNet pack $project --configuration Release --artifacts-path $artifactsPath --no-restore --output $feedPath --property:PackageVersion=$OfficeIMOVersion
    }

    $browserProperties = @(
        "--property:PackageVersion=$OfficeIMOVersion",
        "--property:HtmlTinkerXVersion=$htmlTinkerXVersion"
    )
    Invoke-DotNet restore $browserProjectPath @browserProperties --artifacts-path $artifactsPath --configfile $configPath --packages $packagesPath --no-http-cache --force-evaluate
    Invoke-DotNet pack $browserProjectPath --configuration Release --artifacts-path $artifactsPath --no-restore --output $feedPath @browserProperties

    $consumerProperties = @(
        '--property:EnableOfficeIMOHtmlPdfBrowserPackageSmoke=true',
        "--property:OfficeIMOHtmlPdfBrowserPackageVersion=$OfficeIMOVersion"
    )
    $consumerProjectPath = 'Build/PackageSmoke/OfficeIMO.Html.Pdf.Browser/OfficeIMO.Html.Pdf.Browser.PackageSmoke.csproj'
    Invoke-DotNet restore $consumerProjectPath @consumerProperties --artifacts-path $artifactsPath --configfile $configPath --packages $packagesPath --no-http-cache --force-evaluate
    foreach ($framework in @('net8.0', 'net10.0')) {
        Invoke-DotNet run --project $consumerProjectPath --configuration Release --framework $framework --artifacts-path $artifactsPath --no-restore @consumerProperties
    }

    Write-Host "Validated packed OfficeIMO browser PDF APIs against HtmlTinkerX $htmlTinkerXVersion."
} finally {
    Pop-Location
    if (Test-Path -LiteralPath $workingPath) {
        Remove-Item -LiteralPath $workingPath -Recurse -Force
    }
}
