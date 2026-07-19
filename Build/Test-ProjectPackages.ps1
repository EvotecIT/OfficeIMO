param(
    [Parameter(Mandatory)]
    [string] $FeedPath,

    [string] $Version = '3.0.0',

    [switch] $KeepWorkingDirectory
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

function Get-PackageMetadata {
    param([Parameter(Mandatory)][System.IO.FileInfo] $Package)

    $archive = [System.IO.Compression.ZipFile]::OpenRead($Package.FullName)
    try {
        $nuspecEntry = @($archive.Entries | Where-Object {
                $_.FullName.EndsWith('.nuspec', [StringComparison]::OrdinalIgnoreCase)
            })
        if ($nuspecEntry.Count -ne 1) {
            throw "Package '$($Package.Name)' must contain exactly one nuspec."
        }

        $reader = [System.IO.StreamReader]::new($nuspecEntry[0].Open())
        try {
            [xml] $nuspec = $reader.ReadToEnd()
        } finally {
            $reader.Dispose()
        }

        $namespace = [System.Xml.XmlNamespaceManager]::new($nuspec.NameTable)
        $namespace.AddNamespace('n', $nuspec.DocumentElement.NamespaceURI)
        $metadata = $nuspec.SelectSingleNode('/n:package/n:metadata', $namespace)
        $dependencies = @($metadata.SelectNodes('.//n:dependency', $namespace))
        $readmeEntry = @($archive.Entries | Where-Object {
                $_.FullName.Equals('README.md', [StringComparison]::OrdinalIgnoreCase)
            })

        [pscustomobject] @{
            Id                   = [string] $metadata.id
            Version              = [string] $metadata.version
            Readme               = [string] $metadata.readme
            HasPackagedReadme    = $readmeEntry.Count -eq 1
            DependencyIds        = @($dependencies | ForEach-Object { [string] $_.id })
            OfficeIMODependencies = @($dependencies | Where-Object {
                    ([string] $_.id).StartsWith('OfficeIMO.', [StringComparison]::OrdinalIgnoreCase)
                } | ForEach-Object {
                    [pscustomobject] @{
                        Id      = [string] $_.id
                        Version = [string] $_.version
                    }
                })
        }
    } finally {
        $archive.Dispose()
    }
}

if ($Version -notmatch '^\d+\.\d+\.\d+$') {
    throw "Version must be a public three-part version such as 3.0.0."
}

$resolvedFeed = (Resolve-Path -LiteralPath $FeedPath).Path
$buildConfiguration = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'project.build.json') -Raw |
    ConvertFrom-Json
$packageIds = @($buildConfiguration.ExpectedVersionMap.PSObject.Properties.Name |
        Sort-Object -Unique)
if ($packageIds.Count -eq 0) {
    throw 'Build/project.build.json does not define any coordinated packages.'
}

$workingPath = Join-Path ([System.IO.Path]::GetTempPath()) (
    'officeimo-package-smoke-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $workingPath | Out-Null

try {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $packageMetadata = foreach ($packageId in $packageIds) {
        $matches = @(Get-ChildItem -LiteralPath $resolvedFeed -File -Filter '*.nupkg' |
                Where-Object {
                    $_.BaseName.Equals(
                        "$packageId.$Version",
                        [StringComparison]::OrdinalIgnoreCase)
                })
        if ($matches.Count -ne 1) {
            throw "Expected exactly one $packageId $Version package in '$resolvedFeed'; found $($matches.Count)."
        }

        $metadata = Get-PackageMetadata -Package $matches[0]
        if (!$metadata.Id.Equals($packageId, [StringComparison]::OrdinalIgnoreCase) -or
            !$metadata.Version.Equals($Version, [StringComparison]::Ordinal)) {
            throw "Package identity mismatch in '$($matches[0].Name)'."
        }
        if (!$metadata.Readme.Equals('README.md', [StringComparison]::OrdinalIgnoreCase) -or
            !$metadata.HasPackagedReadme) {
            throw "Package '$packageId' must declare and contain README.md."
        }
        foreach ($dependency in $metadata.OfficeIMODependencies) {
            if ($dependency.Version -notmatch [Regex]::Escape($Version) -or
                $dependency.Version -match '\d+\.\d+\.\d+\.\d+') {
                throw "Package '$packageId' has unaligned OfficeIMO dependency '$($dependency.Id)' version '$($dependency.Version)'."
            }
        }

        $metadata
    }

    $toolPackageId = 'OfficeIMO.Reader.Tool'
    $libraryPackageIds = @($packageIds | Where-Object {
            !$_.Equals($toolPackageId, [StringComparison]::OrdinalIgnoreCase)
        })
    $projectPath = Join-Path $workingPath 'OfficeIMO.ReleaseConsumer.csproj'
    $programPath = Join-Path $workingPath 'Program.cs'
    $nugetConfigPath = Join-Path $workingPath 'NuGet.Config'
    $packagesPath = Join-Path $workingPath 'packages'
    $toolPath = Join-Path $workingPath 'tool'

    $packageReferences = $libraryPackageIds | ForEach-Object {
        '    <PackageReference Include="' +
        [System.Security.SecurityElement]::Escape($_) +
        '" Version="[' + $Version + ']" />'
    }
    $projectXmlLines = @(
        '<Project Sdk="Microsoft.NET.Sdk">',
        '  <PropertyGroup>',
        '    <OutputType>Exe</OutputType>',
        '    <TargetFramework>net8.0</TargetFramework>',
        '    <ImplicitUsings>enable</ImplicitUsings>',
        '    <Nullable>enable</Nullable>',
        '  </PropertyGroup>',
        '  <ItemGroup>'
    ) + $packageReferences + @(
        '  </ItemGroup>',
        '</Project>'
    )
    $projectXml = $projectXmlLines -join [Environment]::NewLine
    Set-Content -LiteralPath $projectPath -Value $projectXml -Encoding utf8
    Set-Content -LiteralPath $programPath -Value (
        'Console.WriteLine("OfficeIMO 3.0 aggregate package consumer loaded.");') -Encoding utf8

    $sourceMappings = $packageIds | ForEach-Object {
        '      <package pattern="' +
        [System.Security.SecurityElement]::Escape($_) +
        '" />'
    }
    $feedXml = [System.Security.SecurityElement]::Escape($resolvedFeed)
    $nugetConfigLines = @(
        '<?xml version="1.0" encoding="utf-8"?>',
        '<configuration>',
        '  <packageSources>',
        '    <clear />',
        "    <add key=`"OfficeIMOLocal`" value=`"$feedXml`" />",
        '    <add key="nuget.org" value="https://api.nuget.org/v3/index.json" protocolVersion="3" />',
        '  </packageSources>',
        '  <packageSourceMapping>',
        '    <packageSource key="OfficeIMOLocal">'
    ) + $sourceMappings + @(
        '    </packageSource>',
        '    <packageSource key="nuget.org">',
        '      <package pattern="*" />',
        '    </packageSource>',
        '  </packageSourceMapping>',
        '</configuration>'
    )
    $nugetConfig = $nugetConfigLines -join [Environment]::NewLine
    Set-Content -LiteralPath $nugetConfigPath -Value $nugetConfig -Encoding utf8

    Invoke-DotNet restore $projectPath --configfile $nugetConfigPath --packages $packagesPath --no-cache --force-evaluate
    Invoke-DotNet build $projectPath --configuration Release --no-restore
    Invoke-DotNet run --project $projectPath --configuration Release --no-build
    Invoke-DotNet tool install $toolPackageId --version $Version --tool-path $toolPath --configfile $nugetConfigPath --no-cache

    $toolExecutableName = if ($IsWindows) {
        'officeimo-reader.exe'
    } else {
        'officeimo-reader'
    }
    $toolExecutable = Join-Path $toolPath $toolExecutableName
    & $toolExecutable capabilities --format json
    if ($LASTEXITCODE -ne 0) {
        throw "The packed $toolPackageId command failed with exit code $LASTEXITCODE."
    }

    Write-Host "Validated $($packageMetadata.Count) coordinated packages at version $Version."
} finally {
    if ($KeepWorkingDirectory) {
        Write-Host "Package-smoke working directory retained at '$workingPath'."
    } elseif (Test-Path -LiteralPath $workingPath) {
        Remove-Item -LiteralPath $workingPath -Recurse -Force
    }
}
