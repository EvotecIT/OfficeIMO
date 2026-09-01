param(
    [Parameter(Mandatory)]
    [string] $FeedPath,

    [string] $Version = '3.3.0',

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

    $expectedDirectDependencies = @{
        'OfficeIMO.Data.Arrow' = @('OfficeIMO.Core')
        'OfficeIMO.Provenance.C2pa' = @('OfficeIMO.Core')
        'OfficeIMO.Security' = @('OfficeIMO.Core')
        'OfficeIMO.IWork' = @('OfficeIMO.Core')
        'OfficeIMO.Word' = @('OfficeIMO.Core', 'OfficeIMO.IWork')
        'OfficeIMO.Excel' = @('OfficeIMO.Core', 'OfficeIMO.IWork')
        'OfficeIMO.PowerPoint' = @('OfficeIMO.Core', 'OfficeIMO.IWork')
        'OfficeIMO.Drawing.HarfBuzz' = @('OfficeIMO.Core')
        'OfficeIMO.Html' = @('OfficeIMO.Core')
        'OfficeIMO.Html.Pdf.Browser' = @('OfficeIMO.Core', 'OfficeIMO.Pdf')
        'OfficeIMO.Html.Rtf' = @('OfficeIMO.Core', 'OfficeIMO.Html', 'OfficeIMO.Rtf')
        'OfficeIMO.Email.Html' = @('OfficeIMO.Email', 'OfficeIMO.Html', 'OfficeIMO.Html.Rtf')
        'OfficeIMO.Mhtml' = @('OfficeIMO.Core', 'OfficeIMO.Email', 'OfficeIMO.Html')
        'OfficeIMO.Email.Image' = @('OfficeIMO.Core', 'OfficeIMO.Email', 'OfficeIMO.Email.Html', 'OfficeIMO.Html')
        'OfficeIMO.Mhtml.Pdf' = @('OfficeIMO.Core', 'OfficeIMO.Html.Pdf', 'OfficeIMO.Mhtml', 'OfficeIMO.Pdf')
        'OfficeIMO.Reader.Html' = @('OfficeIMO.Html', 'OfficeIMO.Markdown.Html', 'OfficeIMO.Reader.Core')
        'OfficeIMO.Reader.Email' = @('OfficeIMO.Email', 'OfficeIMO.Email.Html', 'OfficeIMO.Mhtml', 'OfficeIMO.Reader.Core', 'OfficeIMO.Reader.Html')
        'OfficeIMO.Reader.Epub' = @('OfficeIMO.Epub', 'OfficeIMO.Reader.Core', 'OfficeIMO.Reader.Html')
        'OfficeIMO.Visio.Pdf' = @('OfficeIMO.Core', 'OfficeIMO.Pdf', 'OfficeIMO.Visio')
    }
    foreach ($entry in $expectedDirectDependencies.GetEnumerator()) {
        $metadata = $packageMetadata | Where-Object { $_.Id -eq $entry.Key }
        if ($null -eq $metadata) {
            throw "Dependency contract package '$($entry.Key)' was not packed."
        }
        $actual = @($metadata.OfficeIMODependencies.Id | Sort-Object -Unique)
        $expected = @($entry.Value | Sort-Object -Unique)
        if (($actual -join '|') -ne ($expected -join '|')) {
            throw "Package '$($entry.Key)' OfficeIMO dependencies were '$($actual -join ', ')'; expected '$($expected -join ', ')'."
        }
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
    $visioPdfWorkingPath = Join-Path $workingPath 'visio-pdf'
    New-Item -ItemType Directory -Path $visioPdfWorkingPath | Out-Null
    $visioPdfProjectPath = Join-Path $visioPdfWorkingPath 'OfficeIMO.Visio.Pdf.ReleaseConsumer.csproj'
    $visioPdfProgramPath = Join-Path $visioPdfWorkingPath 'VisioPdfProgram.cs'

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
using Apache.Arrow;
using OfficeIMO.CSV;
using OfficeIMO.Data;
using OfficeIMO.Data.Arrow;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Csv;
using OfficeIMO.Excel.Html;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.PowerPoint.Html;
using OfficeIMO.Visio;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;

internal static class Program
{
    public static async Task Main()
    {
        VerifyHtmlConversionApi();
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
            using (DbDataReader generatedReader = OpenCsv(csvPath))
            {
                ReleaseGeneratedRow generated = generatedReader
                    .RowsAs<ReleaseGeneratedRow>(ReleaseGeneratedRowRowMapping.Configure)
                    .Single();
                if (!string.Equals(generated.Value, "Alpha", StringComparison.Ordinal))
                {
                    throw new InvalidOperationException("Packed OfficeIMO.Data.Generators mapping failed its runtime smoke.");
                }
            }
            using (DbDataReader arrowReader = OpenCsv(csvPath))
            using (RecordBatch arrowBatch = arrowReader
                .ReadArrowBatches(new ArrowReadOptions { BatchSize = 1 })
                .Single())
            {
                if (arrowBatch.Length != 1 || arrowBatch.Schema.FieldsList.Count != 1)
                {
                    throw new InvalidOperationException("Packed OfficeIMO.Data.Arrow adapter failed its runtime smoke.");
                }
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

    private static void VerifyHtmlConversionApi()
    {
        var output = new OfficeHtmlDocumentOptions
        {
            EmitDocumentShell = true,
            IncludeDefaultStyles = true,
            Title = "Package contract",
            Language = "en",
            NewLine = "\n"
        };
        WordToHtmlOptions word = WordToHtmlOptions.CreateDocumentRoundTripProfile();
        word.DocumentOutput = output.Clone();
        ExcelHtmlSaveOptions excel = ExcelHtmlSaveOptions.CreateVisualReviewProfile();
        excel.DocumentOutput = output.Clone();
        PowerPointHtmlSaveOptions powerPoint = PowerPointHtmlSaveOptions.CreateVisualReviewProfile();
        powerPoint.DocumentOutput = output.Clone();
        RtfToHtmlOptions rtf = RtfToHtmlOptions.CreatePrintReviewProfile();
        rtf.DocumentOutput = output.Clone();
        PdfHtmlSaveOptions pdf = PdfHtmlSaveOptions.CreatePositionedReviewProfile();
        pdf.DocumentOutput = output.Clone();

        HtmlTargetCapabilityContract contract = HtmlTargetCapabilityContracts.Get(HtmlConversionTarget.Pdf);
        HtmlToTargetCapabilityContract htmlToTarget = contract.HtmlToTarget;
        TargetToHtmlCapabilityContract targetToHtml = contract.TargetToHtml
            ?? throw new InvalidOperationException("Packed PDF reverse HTML route is missing.");
        if (htmlToTarget.Profiles.Contains("PositionedReview", StringComparer.Ordinal) ||
            !targetToHtml.Profiles.Contains("PositionedReview", StringComparer.Ordinal) ||
            string.IsNullOrWhiteSpace(targetToHtml.DiagnosticsContract))
        {
            throw new InvalidOperationException("Packed directional HTML route contract is inconsistent.");
        }
    }

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

[GenerateRowMapper]
internal sealed class ReleaseGeneratedRow
{
    public string? Value { get; set; }
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

    $visioPdfProjectXml = @"
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <OutputType>Exe</OutputType>
    <TargetFrameworks>net472;net8.0;net10.0</TargetFrameworks>
    <LangVersion>latest</LangVersion>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
    <EnableDefaultCompileItems>false</EnableDefaultCompileItems>
  </PropertyGroup>
  <ItemGroup>
    <Compile Include="VisioPdfProgram.cs" />
    <PackageReference Include="OfficeIMO.Visio.Pdf" Version="[$Version]" />
  </ItemGroup>
</Project>
"@
    Set-Content -LiteralPath $visioPdfProjectPath -Value $visioPdfProjectXml -Encoding utf8
    $visioPdfProgramSource = @"
using OfficeIMO.Visio;
using OfficeIMO.Visio.Pdf;

string[] readerAssemblies = Directory.GetFiles(
    AppContext.BaseDirectory,
    "OfficeIMO.Reader*.dll",
    SearchOption.TopDirectoryOnly);
if (readerAssemblies.Length != 0)
{
    throw new InvalidOperationException(
        "OfficeIMO.Visio.Pdf restored Reader assemblies: " +
        string.Join(", ", readerAssemblies.Select(Path.GetFileName)));
}

string workingPath = Path.Combine(
    Path.GetTempPath(),
    "officeimo-visio-pdf-smoke-" + Guid.NewGuid().ToString("N"));
Directory.CreateDirectory(workingPath);
try
{
    string visioPath = Path.Combine(workingPath, "diagram.vsdx");
    string pdfPath = Path.Combine(workingPath, "diagram.pdf");
    VisioDocument document = VisioDocument.Create(visioPath);
    document.AddPage("Page-1");
    document.Save();

    document.SaveAsPdf(pdfPath).RequireSuccess();
    if (!File.Exists(pdfPath) || new FileInfo(pdfPath).Length == 0)
    {
        throw new InvalidOperationException("Packed OfficeIMO.Visio.Pdf produced no PDF output.");
    }

    Console.WriteLine("OfficeIMO.Visio.Pdf package smoke passed without Reader assemblies.");
}
finally
{
    Directory.Delete(workingPath, recursive: true);
}
"@
    Set-Content -LiteralPath $visioPdfProgramPath -Value $visioPdfProgramSource -Encoding utf8
    Invoke-DotNet restore $visioPdfProjectPath --configfile $nugetConfigPath --packages $packagesPath --no-cache --force-evaluate
    Invoke-DotNet build $visioPdfProjectPath --configuration Release --no-restore
    Invoke-DotNet run --project $visioPdfProjectPath --configuration Release --framework net8.0 --no-build

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
