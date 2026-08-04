param(
    [Parameter(Mandatory)]
    [string] $FeedPath,

    [string] $Version = '3.1.0',

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

    $toolPackages = @(
        [pscustomobject]@{
            Id = 'OfficeIMO.Tool'
            Executable = $(if ($IsWindows) { 'officeimo.exe' } else { 'officeimo' })
            Arguments = @('reader', 'capabilities', '--format', 'json')
        }
    )
    $toolPackageIds = @($toolPackages.Id)
    $libraryPackageIds = @($packageIds | Where-Object {
            $_ -notin $toolPackageIds
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
    $programSource = @"
using System.Data.Common;
using OfficeIMO.CSV;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Csv;
using OfficeIMO.Visio;
using OfficeIMO.Word;

internal static class Program
{
    public static async Task Main()
    {
        string workingPath = Path.Combine(
            Path.GetTempPath(),
            "officeimo-release-api-smoke-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(workingPath);
        try
        {
            string csvPath = Path.Combine(workingPath, "input.csv");
            File.WriteAllText(csvPath, "Value" + Environment.NewLine + "Alpha" + Environment.NewLine);
            using (DbDataReader csvReader = OpenCsv(csvPath))
            {
                if (!csvReader.Read() || !string.Equals(csvReader["Value"]?.ToString(), "Alpha", StringComparison.Ordinal))
                {
                    throw new InvalidOperationException("Packed OfficeIMO.CSV reader failed its runtime smoke.");
                }
            }
            if (!string.Equals(ReadCsvRows(csvPath).Single().Value, "Alpha", StringComparison.Ordinal))
            {
                throw new InvalidOperationException("Packed OfficeIMO.CSV typed mapping failed its runtime smoke.");
            }

            string adapterExcelPath = Path.Combine(workingPath, "csv-import.xlsx");
            CsvDocument csv = CsvDocument.Load(csvPath);
            using (ExcelDocument imported = csv.ToExcelDocument(new ExcelCsvImportOptions
            {
                SheetName = "Data",
                CreateTable = false
            }))
            {
                string roundTripCsv = imported["Data"].ToCsv();
                if (!roundTripCsv.Contains("Alpha", StringComparison.Ordinal))
                {
                    throw new InvalidOperationException("Packed OfficeIMO.Excel.Csv round trip failed its runtime smoke.");
                }
                imported.Save(adapterExcelPath);
            }

            string excelPath = Path.Combine(workingPath, "report.xlsx");
            using (ExcelDocument excel = CreateExcel(excelPath))
            {
                excel.Save();
            }
            using (DbDataReader excelReader = OpenExcel(excelPath))
            {
                if (!excelReader.Read() || !string.Equals(excelReader["Value"]?.ToString(), "Alpha", StringComparison.Ordinal))
                {
                    throw new InvalidOperationException("Packed OfficeIMO.Excel reader failed its runtime smoke.");
                }
            }

            string wordPath = Path.Combine(workingPath, "report.docx");
            using (WordDocument word = CreateWord(wordPath))
            {
                word.Save();
            }
            using (WordDocument loadedWord = WordDocument.Load(wordPath))
            {
                if (!loadedWord.Paragraphs.Any(paragraph => paragraph.Text.Contains("OfficeIMO.Word fluent API.", StringComparison.Ordinal)))
                {
                    throw new InvalidOperationException("Packed OfficeIMO.Word fluent API failed its runtime smoke.");
                }
            }

            string visioPath = Path.Combine(workingPath, "diagram.vsdx");
            VisioDocument visio = VisioDocument.Create(visioPath);
            visio.AddPage("Page-1");
            visio.Save();
            VisioDocument loadedVisio = await LoadVisioAsync(visioPath, options: null, cancellationToken: default);
            if (loadedVisio.Pages.Count != 1)
            {
                throw new InvalidOperationException("Packed OfficeIMO.Visio async load failed its runtime smoke.");
            }

            Console.WriteLine("OfficeIMO $Version aggregate package runtime smoke passed.");
        }
        finally
        {
            Directory.Delete(workingPath, recursive: true);
        }
    }

    private static DbDataReader OpenCsv(string path) =>
        CsvDocument.OpenDataReader(path);

    private static IEnumerable<ReleaseRow> ReadCsvRows(string path) =>
        CsvDocument.Load(path).RowsAs<ReleaseRow>();

    private static DbDataReader OpenExcel(string path) =>
        ExcelDocument.OpenDataReader(path, new ExcelReadOptions { SheetName = "Data" });

    private static IEnumerable<ReleaseRow> ReadExcelRows(ExcelSheet sheet) =>
        sheet.RowsAs<ReleaseRow>();

    private static ExcelDocument CreateExcel(string path)
    {
        ExcelDocument document = ExcelDocument.Create(path);
        return document.AsFluent()
            .Sheet("Data", sheet => sheet
                .Cell(1, 1, "Value")
                .Cell(2, 1, "Alpha"))
            .End();
    }

    private static WordDocument CreateWord(string path)
    {
        WordDocument document = WordDocument.Create(path);
        return document.AsFluent()
            .H1("Package smoke")
            .Paragraph(paragraph => paragraph.Text("OfficeIMO.Word fluent API."))
            .End();
    }

    private static Task<VisioDocument> LoadVisioAsync(
        string path,
        VisioLoadOptions? options,
        CancellationToken cancellationToken) =>
        VisioDocument.LoadAsync(path, options, cancellationToken);

    private sealed class ReleaseRow
    {
        public string? Value { get; set; }
    }
}
"@
    Set-Content -LiteralPath $programPath -Value $programSource -Encoding utf8

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
    foreach ($toolPackage in $toolPackages) {
        Invoke-DotNet tool install $toolPackage.Id --version $Version --tool-path $toolPath --configfile $nugetConfigPath --no-cache
        $toolExecutable = Join-Path $toolPath $toolPackage.Executable
        $toolArguments = @($toolPackage.Arguments)
        & $toolExecutable @toolArguments
        if ($LASTEXITCODE -ne 0) {
            throw "The packed $($toolPackage.Id) command failed with exit code $LASTEXITCODE."
        }
    }

    Write-Host "Validated $($packageMetadata.Count) coordinated packages at version $Version."
} finally {
    if ($KeepWorkingDirectory) {
        Write-Host "Package-smoke working directory retained at '$workingPath'."
    } elseif (Test-Path -LiteralPath $workingPath) {
        Remove-Item -LiteralPath $workingPath -Recurse -Force
    }
}
