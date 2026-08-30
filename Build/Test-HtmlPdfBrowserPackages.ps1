param(
    [string] $HtmlTinkerXPackagePath,
    [string] $OfficeIMOVersion = '3.2.5-browser-local',
    [version] $MinimumHtmlTinkerXVersion = '3.0.1',
    [string] $ExpectedHtmlTinkerXCommit
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

function ConvertFrom-NuGetNuspec {
    param([Parameter(Mandatory)][xml] $Nuspec)

    $namespace = [System.Xml.XmlNamespaceManager]::new($Nuspec.NameTable)
    $namespace.AddNamespace('n', $Nuspec.DocumentElement.NamespaceURI)
    $metadata = $Nuspec.SelectSingleNode('/n:package/n:metadata', $namespace)
    [pscustomobject]@{
        Id = [string] $metadata.id
        Version = [version] ([string] $metadata.version)
        RepositoryUrl = [string] $metadata.repository.url
        RepositoryType = [string] $metadata.repository.type
        RepositoryCommit = [string] $metadata.repository.commit
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
        ConvertFrom-NuGetNuspec -Nuspec $nuspec
    } finally {
        $archive.Dispose()
    }
}

function Assert-HtmlTinkerXIdentity {
    param([Parameter(Mandatory)][object] $Identity)

    if (-not $Identity.Id.Equals('HtmlTinkerX', [StringComparison]::OrdinalIgnoreCase)) {
        throw "Selected package is '$($Identity.Id)', not HtmlTinkerX."
    }
    if (-not [string]::IsNullOrWhiteSpace($ExpectedHtmlTinkerXCommit)) {
        if ($ExpectedHtmlTinkerXCommit -notmatch '^[0-9a-fA-F]{40}$') {
            throw 'ExpectedHtmlTinkerXCommit must be a full 40-character Git commit.'
        }
        if ([string] $Identity.RepositoryType -ne 'git' -or
            [string] $Identity.RepositoryUrl -notmatch '(?i)^https://github\.com/EvotecIT/HtmlTinkerX(?:\.git)?$' -or
            -not [string]::Equals(
                [string] $Identity.RepositoryCommit,
                $ExpectedHtmlTinkerXCommit,
                [StringComparison]::OrdinalIgnoreCase)) {
            throw "HtmlTinkerX package provenance does not match measured commit $ExpectedHtmlTinkerXCommit."
        }
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
    Assert-HtmlTinkerXIdentity -Identity $identity
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
    $installedNuspecPath = Join-Path $packagesPath (
        "htmltinkerx/$($htmlTinkerXVersion.ToString())/htmltinkerx.nuspec")
    if (-not (Test-Path -LiteralPath $installedNuspecPath -PathType Leaf)) {
        throw "Restored HtmlTinkerX nuspec was not found at '$installedNuspecPath'."
    }
    [xml] $installedNuspec = Get-Content -LiteralPath $installedNuspecPath -Raw
    Assert-HtmlTinkerXIdentity -Identity (ConvertFrom-NuGetNuspec -Nuspec $installedNuspec)
    Invoke-DotNet pack $browserProjectPath --configuration Release --artifacts-path $artifactsPath --no-restore --output $feedPath @browserProperties

    $consumerProperties = @(
        '--property:EnableOfficeIMOHtmlPdfBrowserPackageSmoke=true',
        "--property:OfficeIMOHtmlPdfBrowserPackageVersion=$OfficeIMOVersion"
    )
    $consumerProjectPath = 'Build/PackageSmoke/OfficeIMO.Html.Pdf.Browser/OfficeIMO.Html.Pdf.Browser.PackageSmoke.csproj'
    Invoke-DotNet restore $consumerProjectPath @consumerProperties --artifacts-path $artifactsPath --configfile $configPath --packages $packagesPath --no-http-cache --force-evaluate
    $frameworks = if ($IsWindows) { @('net472', 'net8.0', 'net10.0') } else { @('net8.0', 'net10.0') }
    foreach ($framework in $frameworks) {
        Invoke-DotNet run --project $consumerProjectPath --configuration Release --framework $framework --artifacts-path $artifactsPath --no-restore @consumerProperties
    }

    Write-Host "Validated packed OfficeIMO browser PDF APIs against HtmlTinkerX $htmlTinkerXVersion."
} finally {
    Pop-Location
    if (Test-Path -LiteralPath $workingPath) {
        Remove-Item -LiteralPath $workingPath -Recurse -Force
    }
}
