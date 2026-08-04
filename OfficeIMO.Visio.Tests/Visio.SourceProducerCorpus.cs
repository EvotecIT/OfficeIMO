using System.Security.Cryptography;
using System.IO.Compression;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Visio;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class VisioSourceProducerCorpusTests {
    private static string CorpusDirectory => Path.Combine(GetRepositoryRoot(),
        "Assets", "VisioTemplates");

    public static IEnumerable<object[]> CorpusArtifacts() {
        VisioSourceCorpusManifest manifest = LoadManifest();
        foreach (VisioSourceCorpusArtifact artifact in manifest.Artifacts) {
            yield return new object[] {
                artifact.File,
                (VisioPackageType)Enum.Parse(typeof(VisioPackageType),
                    artifact.PackageType, ignoreCase: false),
                artifact.Sha256
            };
        }
    }

    [Theory]
    [MemberData(nameof(CorpusArtifacts))]
    public void MicrosoftVisioPackageFamiliesLoadReadAndRoundTrip(
        string file, VisioPackageType packageType, string sha256) {
        string source = Path.Combine(CorpusDirectory, file);
        string roundTrip = Path.Combine(Path.GetTempPath(),
            "OfficeIMO-VisioProducer-" + Guid.NewGuid().ToString("N")
            + Path.GetExtension(file));
        try {
            Assert.Equal(sha256, ComputeSha256(source));
            Assert.Empty(VisioValidator.Validate(source));
            VisioDocument loaded = VisioDocument.Load(source);
            Assert.Equal(packageType, loaded.PackageType);
            VisioPage page = Assert.Single(loaded.Pages);
            Assert.Equal("Advanced Corpus", page.Name);
            Assert.True(page.Shapes.Count >= 16);
            Assert.Equal(11, page.Connectors.Count);
            Assert.Contains(page.Shapes, shape => shape.Children.Count == 2);
            Assert.Contains(page.Shapes, shape => shape.IsContainer);
            VisioShape firstStep = Assert.Single(page.Shapes,
                shape => shape.Text?.Trim() == "Step 1.1");
            Assert.Equal("Team 1",
                firstStep.GetShapeDataValue("Owner"));
            Assert.Equal("Healthy", firstStep.GetUserCellValue("Status"));

            Assert.Single(loaded.ToOfficeDocumentReadResult().Pages);
            OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
                .AddVisioHandler().Build();
            Assert.Single(reader.ReadDocument(source).Pages);

            loaded.Save(roundTrip, packageType);
            Assert.Empty(VisioValidator.Validate(roundTrip));
            VisioDocument reopened = VisioDocument.Load(roundTrip);
            Assert.Equal(packageType, reopened.PackageType);
            VisioPage reopenedPage = Assert.Single(reopened.Pages);
            Assert.Equal(11, reopenedPage.Connectors.Count);
            Assert.Contains(reopenedPage.Shapes,
                shape => shape.Children.Count == 2);
            Assert.Contains(reopenedPage.Shapes, shape => shape.IsContainer);
            AssertShapeDataSectionName(roundTrip, "Property");
            AssertConnectorGeometry(roundTrip, expectedCount: 11);
            AssertDynamicConnectorMaster(roundTrip, expectedCount: 11);
        } finally {
            if (File.Exists(roundTrip)) File.Delete(roundTrip);
        }
    }

    [Fact]
    public void MicrosoftVisioDrawingSupportsTypedEditingRelayoutAndExports() {
        string source = Path.Combine(CorpusDirectory,
            "VisioAdvancedRoadmap.vsdx");
        string roundTrip = Path.Combine(Path.GetTempPath(),
            "OfficeIMO-VisioAdvancedEdit-" + Guid.NewGuid().ToString("N")
            + ".vsdx");
        string pdf = Path.ChangeExtension(roundTrip, ".pdf");
        try {
            VisioDocument document = VisioDocument.Load(source);
            VisioPage page = Assert.Single(document.Pages);
            VisioShape target = Assert.Single(page.Shapes,
                shape => shape.Text?.Trim() == "Step 1.1");
            target.SetShapeData("Slo", "91");
            VisioDataGraphic graphic = VisioDataGraphic.Create()
                .Badge("Status")
                .Bar("Slo", maximumValue: 100, label: "SLO");
            Assert.NotEmpty(page.AddDataGraphics(target, graphic));
            Assert.Equal(4, page.AddDataGraphicLegend(
                "corpus-legend", "Imported health", graphic, 11.7, 8.2)
                .Shapes.Count);

            VisioShape inner = page.AddContainer("typed-inner",
                "Typed inner", new[] { target });
            VisioShape outer = page.AddContainer("typed-outer",
                "Typed outer", new[] { inner });
            page.AddNestedContainer(outer, inner);
            Assert.Equal(1, page.GetContainerInfo(inner).NestingDepth);

            VisioComment root = page.AddComment(target,
                "Review imported step", "Owner", "OW");
            page.ReplyToComment(root.Id, "Reviewed",
                new VisioCommentAuthor("Reviewer", "RV",
                    "reviewer@example.test"));
            Assert.Equal(2,
                Assert.Single(page.GetCommentThreads()).Comments.Count);

            page.RelayoutDiagram(new VisioWholeDiagramRelayoutOptions {
                PolishAfterLayout = true,
                RouteConnectors = true
            });
            page.FitToContent(new VisioFitToContentOptions {
                IncludeGroupChildren = true,
                IncludeConnectors = true,
                HorizontalMargin = 0.3,
                VerticalMargin = 0.3
            });

            OfficeImageExportResult png = page.ExportImage(
                OfficeImageExportFormat.Png);
            Assert.DoesNotContain(png.Diagnostics, diagnostic =>
                diagnostic.Severity ==
                OfficeImageExportDiagnosticSeverity.Error);
            Assert.True(OfficePngReader.TryDecode(png.Bytes,
                out OfficeRasterImage? raster));
            Assert.True(raster!.Width > 0 && raster.Height > 0);
            document.SaveAsPdf(pdf);
            Assert.Equal("%PDF-", System.Text.Encoding.ASCII.GetString(
                File.ReadAllBytes(pdf), 0, 5));
            document.Save(roundTrip);

            VisioDocument reopened = VisioDocument.Load(roundTrip);
            VisioPage reopenedPage = Assert.Single(reopened.Pages);
            VisioShape reopenedTarget = Assert.Single(reopenedPage.Shapes,
                shape => shape.Text?.Trim() == "Step 1.1");
            Assert.NotEmpty(reopenedPage.GetDataGraphic(reopenedTarget).Shapes);
            Assert.Equal(2,
                Assert.Single(reopenedPage.GetCommentThreads()).Comments.Count);
            Assert.Equal(1, reopenedPage.GetContainerInfo(
                Assert.Single(reopenedPage.Shapes,
                    shape => shape.Id == "typed-inner")).NestingDepth);
        } finally {
            if (File.Exists(roundTrip)) File.Delete(roundTrip);
            if (File.Exists(pdf)) File.Delete(pdf);
        }
    }

    [Fact]
    public void ValidatorRejectsUnsupportedPackageExtension() {
        string source = Path.Combine(CorpusDirectory,
            "VisioAdvancedRoadmap.vsdx");
        string renamed = Path.Combine(Path.GetTempPath(),
            "OfficeIMO-VisioUnsupported-" + Guid.NewGuid().ToString("N")
            + ".foo");
        try {
            File.Copy(source, renamed);
            Assert.Contains(VisioValidator.Validate(renamed), issue =>
                issue.Contains("not a supported Visio Open XML package extension",
                    StringComparison.Ordinal));
        } finally {
            if (File.Exists(renamed)) File.Delete(renamed);
        }
    }

    [Fact]
    public void CreateRejectsUnsupportedPackageExtension() {
        string path = Path.Combine(Path.GetTempPath(),
            "OfficeIMO-VisioUnsupported-" + Guid.NewGuid().ToString("N")
            + ".foo");
        ArgumentException exception = Assert.Throws<ArgumentException>(() =>
            VisioDocument.Create(path));
        Assert.Contains("supported Visio Open XML package extension",
            exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    [Trait("Category", "MicrosoftOfficeInteroperability")]
    public void DesktopVisioOpensRoundTripsAndExportsCorpusWhenRequested() {
        if (!string.Equals(Environment.GetEnvironmentVariable(
                "OFFICEIMO_RUN_VISIO_DESKTOP_CORPUS"), "1",
                StringComparison.Ordinal)) return;

        string? configuredOutput = Environment.GetEnvironmentVariable(
            "OFFICEIMO_VISIO_DESKTOP_CORPUS_OUTPUT");
        string output = string.IsNullOrWhiteSpace(configuredOutput)
            ? Path.Combine(Path.GetTempPath(),
                "OfficeIMO-VisioProducerDesktop-"
                + Guid.NewGuid().ToString("N"))
            : Path.GetFullPath(configuredOutput);
        bool deleteOutput = string.IsNullOrWhiteSpace(configuredOutput);
        Directory.CreateDirectory(output);
        try {
            foreach (VisioSourceCorpusArtifact artifact in
                     LoadManifest().Artifacts) {
                string source = Path.Combine(CorpusDirectory, artifact.File);
                string prefix = Path.GetFileNameWithoutExtension(artifact.File);
                string officeImoRoundTrip = Path.Combine(output,
                    prefix + "-officeimo" + Path.GetExtension(artifact.File));
                VisioDocument officeImoDocument = VisioDocument.Load(source);
                VisioPackageType packageType = (VisioPackageType)Enum.Parse(
                    typeof(VisioPackageType), artifact.PackageType,
                    ignoreCase: false);
                officeImoDocument.Save(officeImoRoundTrip, packageType);
                var options = new VisioDesktopValidationOptions {
                    SaveCopy = true,
                    SaveCopyPath = Path.Combine(output,
                        prefix + "-visio-roundtrip"
                        + Path.GetExtension(artifact.File)),
                    ExportDirectory = output,
                    ExportFileNamePrefix = prefix + "-officeimo"
                };
                if (string.Equals(Path.GetExtension(artifact.File), ".vsdx",
                        StringComparison.OrdinalIgnoreCase)) {
                    options.ExportFormats.Add(VisioDesktopExportFormat.Png);
                    options.ExportFormats.Add(VisioDesktopExportFormat.Svg);
                    options.ExportFormats.Add(VisioDesktopExportFormat.Pdf);
                }
                VisioDesktopValidationResult result =
                    VisioDesktopBaselineValidator.Validate(officeImoRoundTrip,
                        options);
                Assert.True(result.IsAvailable,
                    string.Join(Environment.NewLine, result.Issues));
                Assert.True(result.IsValid,
                    string.Join(Environment.NewLine, result.Issues));
                Assert.Contains(options.SaveCopyPath, result.OutputFiles);
                AssertShapeDataSectionName(officeImoRoundTrip, "Property");
                AssertDynamicConnectorMaster(officeImoRoundTrip,
                    expectedCount: 11);
                if (string.Equals(Path.GetExtension(artifact.File), ".vsdx",
                        StringComparison.OrdinalIgnoreCase)) {
                    string svg = Assert.Single(result.OutputFiles,
                        path => string.Equals(Path.GetExtension(path), ".svg",
                            StringComparison.OrdinalIgnoreCase));
                    XDocument svgDocument = XDocument.Load(svg);
                    string svgStyles = string.Concat(svgDocument.Descendants()
                        .Where(element => string.Equals(element.Name.LocalName,
                            "style", StringComparison.OrdinalIgnoreCase))
                        .Select(element => element.Value));
                    HashSet<string> arrowMarkerClasses = Regex.Matches(svgStyles,
                            @"\.(?<class>[A-Za-z_][\w-]*)\s*\{[^}]*marker-end\s*:",
                            RegexOptions.IgnoreCase)
                        .Cast<Match>()
                        .Select(match => match.Groups["class"].Value)
                        .ToHashSet(StringComparer.Ordinal);
                    int renderedArrowheads = svgDocument.Descendants()
                        .Where(element => string.Equals(element.Name.LocalName,
                            "path", StringComparison.OrdinalIgnoreCase))
                        .SelectMany(element => ((string?)element.Attribute("class")
                                ?? string.Empty).Split(new[] { ' ' },
                                StringSplitOptions.RemoveEmptyEntries))
                        .Count(arrowMarkerClasses.Contains);
                    Assert.True(renderedArrowheads >= 11,
                        $"Expected at least 11 rendered connector arrowheads, but found {renderedArrowheads}.");
                }
            }
        } finally {
            if (deleteOutput) {
                try { Directory.Delete(output, recursive: true); } catch { }
            }
        }
    }

    private static VisioSourceCorpusManifest LoadManifest() {
        string path = Path.Combine(CorpusDirectory, "corpus-manifest.json");
        VisioSourceCorpusManifest manifest = JsonSerializer.Deserialize<
            VisioSourceCorpusManifest>(File.ReadAllText(path),
            new JsonSerializerOptions { PropertyNameCaseInsensitive = true })
            ?? throw new InvalidDataException("Visio source corpus manifest is empty.");
        Assert.Equal(1, manifest.SchemaVersion);
        Assert.Equal("Microsoft Visio 16", manifest.Producer);
        Assert.Equal(6, manifest.Artifacts.Count);
        string[] declared = manifest.Artifacts.Select(artifact => artifact.File)
            .OrderBy(file => file, StringComparer.OrdinalIgnoreCase).ToArray();
        string[] actual = Directory.GetFiles(CorpusDirectory,
                "VisioAdvancedRoadmap.*")
            .Select(Path.GetFileName)
            .Where(file => file != null)
            .OrderBy(file => file, StringComparer.OrdinalIgnoreCase)
            .ToArray()!;
        Assert.Equal(declared, actual, StringComparer.OrdinalIgnoreCase);
        return manifest;
    }

    private static string ComputeSha256(string path) {
        using SHA256 sha256 = SHA256.Create();
        using FileStream stream = File.OpenRead(path);
        return string.Concat(sha256.ComputeHash(stream)
            .Select(value => value.ToString("x2")));
    }

    private static void AssertShapeDataSectionName(string path,
        string expectedName) {
        using ZipArchive archive = ZipFile.OpenRead(path);
        ZipArchiveEntry pageEntry = archive.GetEntry("visio/pages/page1.xml")
            ?? throw new InvalidDataException("The package has no first Visio page.");
        using Stream stream = pageEntry.Open();
        XDocument pageXml = XDocument.Load(stream);
        string[] names = pageXml.Descendants()
            .Where(element => element.Name.LocalName == "Section")
            .Select(element => (string?)element.Attribute("N"))
            .Where(name => name != null)
            .Select(name => name!)
            .Where(name => name is "Prop" or "Property")
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
        Assert.Equal(new[] { expectedName }, names,
            StringComparer.OrdinalIgnoreCase);
    }

    private static void AssertConnectorGeometry(string path,
        int expectedCount) {
        using ZipArchive archive = ZipFile.OpenRead(path);
        ZipArchiveEntry pageEntry = archive.GetEntry("visio/pages/page1.xml")
            ?? throw new InvalidDataException(
                "The package has no first Visio page.");
        using Stream stream = pageEntry.Open();
        XDocument pageXml = XDocument.Load(stream);
        int count = pageXml.Descendants()
            .Where(element => element.Name.LocalName == "Shape")
            .Count(shape => shape.Elements().Any(section =>
                section.Name.LocalName == "Section"
                && string.Equals((string?)section.Attribute("N"),
                    "Geometry", StringComparison.OrdinalIgnoreCase))
                && shape.Elements().Any(cell =>
                    cell.Name.LocalName == "Cell"
                    && (string?)cell.Attribute("N") == "BeginX"));
        Assert.Equal(expectedCount, count);
    }

    private static void AssertDynamicConnectorMaster(string path,
        int expectedCount) {
        using ZipArchive archive = ZipFile.OpenRead(path);
        Assert.NotNull(archive.GetEntry("visio/masters/masters.xml"));
        ZipArchiveEntry masterEntry = Assert.Single(archive.Entries,
            entry => entry.FullName.StartsWith("visio/masters/master",
                         StringComparison.OrdinalIgnoreCase)
                     && !string.Equals(entry.FullName,
                         "visio/masters/masters.xml",
                         StringComparison.OrdinalIgnoreCase)
                     && entry.FullName.EndsWith(".xml",
                         StringComparison.OrdinalIgnoreCase));
        using (Stream masterStream = masterEntry.Open()) {
            XDocument masterXml = XDocument.Load(masterStream);
            XElement masterShape = Assert.Single(masterXml.Descendants(),
                element => element.Name.LocalName == "Shape");
            Assert.Contains(masterShape.Elements(), cell =>
                cell.Name.LocalName == "Cell"
                && (string?)cell.Attribute("N") == "BeginX");
            Assert.Contains(masterShape.Elements(), cell =>
                cell.Name.LocalName == "Cell"
                && (string?)cell.Attribute("N") == "EndX");
            XElement geometry = Assert.Single(masterShape.Elements(),
                element => element.Name.LocalName == "Section"
                           && string.Equals(
                               (string?)element.Attribute("N"),
                               "Geometry",
                               StringComparison.OrdinalIgnoreCase));
            Assert.Contains(geometry.Elements(),
                element => element.Name.LocalName == "Row");
        }
        ZipArchiveEntry pageEntry = archive.GetEntry("visio/pages/page1.xml")
            ?? throw new InvalidDataException(
                "The package has no first Visio page.");
        using Stream stream = pageEntry.Open();
        XDocument pageXml = XDocument.Load(stream);
        int count = pageXml.Descendants()
            .Where(element => element.Name.LocalName == "Shape")
            .Count(shape => shape.Attribute("Master") != null &&
                shape.Elements().Any(cell =>
                    cell.Name.LocalName == "Cell" &&
                    (string?)cell.Attribute("N") == "BeginX"));
        Assert.Equal(expectedCount, count);
    }

    private static string GetRepositoryRoot() {
        DirectoryInfo? directory = new(AppContext.BaseDirectory);
        while (directory != null) {
            if (File.Exists(Path.Combine(directory.FullName, "OfficeIMO.sln")))
                return directory.FullName;
            directory = directory.Parent;
        }
        throw new DirectoryNotFoundException(
            "Could not locate the OfficeIMO repository root.");
    }

    private sealed class VisioSourceCorpusManifest {
        public int SchemaVersion { get; set; }
        public string Producer { get; set; } = string.Empty;
        public List<VisioSourceCorpusArtifact> Artifacts { get; set; } = new();
    }

    private sealed class VisioSourceCorpusArtifact {
        public string File { get; set; } = string.Empty;
        public string Sha256 { get; set; } = string.Empty;
        public string PackageType { get; set; } = string.Empty;
    }
}
