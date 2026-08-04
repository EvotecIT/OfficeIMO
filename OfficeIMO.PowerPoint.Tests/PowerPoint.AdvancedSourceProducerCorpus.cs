using System.Security.Cryptography;
using System.IO.Compression;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

[Collection(PowerPointNonParallelCollection.Name)]
public sealed class PowerPointAdvancedSourceProducerCorpusTests {
    private static string FixturePath => Path.Combine(GetRepositoryRoot(),
        "Assets", "PowerPointTemplates", "PowerPointAdvancedRoadmap.pptx");

    [Fact]
    public void MicrosoftAdvancedCorpusEditsAndRendersWithoutFlatteningNativeContent() {
        string roundTrip = Path.Combine(Path.GetTempPath(),
            "OfficeIMO-PowerPointAdvancedCorpus-" + Guid.NewGuid().ToString("N")
            + ".pptx");
        try {
            using (PowerPointPresentation presentation =
                   PowerPointPresentation.Load(FixturePath)) {
                AssertMicrosoftPowerPointProducer(presentation);
                Assert.Equal(11, presentation.Slides.Count);

                PowerPointChart[] charts = presentation.Slides
                    .SelectMany(slide => slide.Charts).ToArray();
                Assert.Equal(8, charts.Length);
                string[] expectedFamilies = {
                    "area3DChart", "bar3DChart", "line3DChart", "ofPieChart",
                    "pie3DChart", "stockChart", "surface3DChart", "surfaceChart"
                };
                PowerPointImportedChartReport[] reports = charts
                    .Select(chart => chart.InspectImportedContent()).ToArray();
                Assert.Equal(expectedFamilies, reports.Select(report => report.Family)
                    .OrderBy(value => value, StringComparer.Ordinal).ToArray());
                PowerPointImportedChartReport[] editable = reports.Where(
                    report => report.Support ==
                        PowerPointImportedChartSupport.EditableWithProjectedRendering)
                    .ToArray();
                Assert.Equal(6, editable.Length);
                Assert.Equal(new[] { "area3DChart", "stockChart" }, reports
                    .Where(report => report.Support ==
                        PowerPointImportedChartSupport.PreservationOnly)
                    .Select(report => report.Family)
                    .OrderBy(value => value, StringComparer.Ordinal).ToArray());
                Assert.All(reports.Where(report => report.Support ==
                    PowerPointImportedChartSupport.PreservationOnly), report =>
                    Assert.Contains("numeric category storage",
                        report.Diagnostics.Single(),
                        StringComparison.OrdinalIgnoreCase));

                Dictionary<string, string> workbookShellBefore =
                    CaptureBar3DWorkbookShell(presentation);
                ApplyRepresentativeTypedEdits(presentation);
                Assert.Equal(workbookShellBefore.OrderBy(pair => pair.Key),
                    CaptureBar3DWorkbookShell(presentation)
                        .OrderBy(pair => pair.Key));

                PowerPointAnimationReport animations =
                    presentation.InspectAnimations();
                Assert.True(animations.HasAnimations);
                Assert.NotEmpty(animations.Nodes);

                foreach (PowerPointSlide slide in presentation.Slides
                             .Where(slide => slide.Charts.Any())) {
                    OfficeImageExportResult image = slide.ExportImage(
                        OfficeImageExportFormat.Png);
                    Assert.DoesNotContain(image.Diagnostics, diagnostic =>
                        diagnostic.Severity ==
                        OfficeImageExportDiagnosticSeverity.Error);
                    Assert.True(OfficePngReader.TryDecode(image.Bytes,
                        out OfficeRasterImage? raster));
                    Assert.True(raster!.Width > 0 && raster.Height > 0);
                }
                byte[] pdf = presentation.ToPdf();
                Assert.True(pdf.Length > 100);
                Assert.Equal("%PDF-", System.Text.Encoding.ASCII.GetString(
                    pdf, 0, 5));
                presentation.Save(roundTrip);
            }

            using PowerPointPresentation reopened =
                PowerPointPresentation.Load(roundTrip);
            PowerPointChart reopenedChart = reopened.Slides
                .SelectMany(slide => slide.Charts).Single(chart =>
                    chart.InspectImportedContent().Family == "bar3DChart");
            Assert.Equal("bar3DChart",
                reopenedChart.InspectImportedContent().Family);
            Assert.True(reopenedChart.TryGetOfficeSnapshot(
                out OfficeChartSnapshot reopenedSnapshot));
            Assert.Contains(reopenedSnapshot.Data.Series, series =>
                series.Name == "OfficeIMO edited bar3DChart");
            Assert.All(reopened.Slides.SelectMany(slide => slide.Charts)
                    .Where(chart => chart.InspectImportedContent().Support ==
                        PowerPointImportedChartSupport
                            .EditableWithProjectedRendering),
                chart => {
                    Assert.True(chart.TryGetOfficeSnapshot(
                        out OfficeChartSnapshot snapshot));
                    Assert.StartsWith("OfficeIMO edited ",
                        snapshot.Data.Series[0].Name,
                        StringComparison.Ordinal);
                });
            Assert.Equal(new[] { "Validate", "Round-trip", "Represent", "Discover" },
                Assert.Single(reopened.Slides.SelectMany(slide => slide.SmartArts))
                    .GetNodeTexts());
            Assert.Equal(new[] { "Validate", "Round-trip", "Represent", "Discover" },
                GetPersistedSmartArtNodeTextsByHorizontalPosition(reopened));
            PowerPointMedia reopenedAudio = Assert.Single(reopened.Slides
                .SelectMany(slide => slide.Media));
            PowerPointMediaPlaybackOptions playback =
                reopenedAudio.GetPlaybackOptions();
            Assert.Equal(65, playback.VolumePercent);
            Assert.True(playback.Loop);
            Assert.Equal(50U, playback.TrimStartMilliseconds);
            Assert.Equal(900U, playback.TrimEndMilliseconds);
        } finally {
            if (File.Exists(roundTrip)) File.Delete(roundTrip);
        }
    }

    [Fact]
    public void ImportedAdvancedChartPreservesWorkbookAndCacheNumberFormats() {
        using PowerPointPresentation presentation =
            PowerPointPresentation.Load(FixturePath);
        PowerPointChart chart = GetBar3DChart(presentation);
        ChartPart chartPart = GetBar3DChartPart(presentation);
        EmbeddedPackagePart package = Assert.Single(
            chartPart.GetPartsOfType<EmbeddedPackagePart>());
        SetWorkbookCellStyle(package, "B2", styleIndex: 1U);
        NumberingCache cache = chartPart.ChartSpace
            .Descendants<Bar3DChart>().Single()
            .Descendants<NumberingCache>().First();
        cache.FormatCode = new FormatCode("$#,##0.00");

        Assert.True(chart.TryGetOfficeSnapshot(
            out OfficeChartSnapshot snapshot));
        chart.UpdateData(snapshot.Data);

        Assert.Equal(1U, GetWorkbookCellStyle(package, "B2"));
        Assert.Equal("$#,##0.00", chartPart.ChartSpace
            .Descendants<Bar3DChart>().Single()
            .Descendants<NumberingCache>().First()
            .FormatCode?.Text);
    }

    [Fact]
    public void FailedImportedAdvancedChartUpdateLeavesChartAndWorkbookUnchanged() {
        using PowerPointPresentation presentation =
            PowerPointPresentation.Load(FixturePath);
        PowerPointChart chart = GetBar3DChart(presentation);
        ChartPart chartPart = GetBar3DChartPart(presentation);
        EmbeddedPackagePart package = Assert.Single(
            chartPart.GetPartsOfType<EmbeddedPackagePart>());
        RemoveLastWorkbookTableColumn(package);
        string chartBefore = chartPart.ChartSpace.OuterXml;
        string workbookBefore = ComputePartSha256(package);
        Assert.True(chart.TryGetOfficeSnapshot(
            out OfficeChartSnapshot snapshot));

        Assert.Throws<NotSupportedException>(() =>
            chart.UpdateData(snapshot.Data));

        Assert.Equal(chartBefore, chartPart.ChartSpace.OuterXml);
        Assert.Equal(workbookBefore, ComputePartSha256(package));
    }

    [Fact]
    public void SharedChartValidationRejectsUnsafeWorkbookData() {
        Assert.Throws<ArgumentException>(() =>
            PowerPointUtils.ValidateSharedWorkbookDimensions(
                categoryCount: 1_048_576, seriesCount: 1,
                totalPoints: 1));
        Assert.Throws<ArgumentException>(() =>
            PowerPointUtils.ValidateSharedWorkbookDimensions(
                categoryCount: 1, seriesCount: 16_384,
                totalPoints: 1));
        Assert.Throws<ArgumentException>(() =>
            PowerPointUtils.ValidateSharedWorkbookDimensions(
                categoryCount: 1, seriesCount: 1,
                totalPoints: 100_001));
        OfficeChartData nonFinite = new(new[] { "A" },
            new[] { new OfficeChartSeries("Series", new[] { double.NaN }) });
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            PowerPointUtils.ValidateSharedChartData(nonFinite,
                OfficeChartKind.ColumnClustered));
        OfficeChartData nonFiniteScatterCategory = new(
            new[] { "NaN" },
            new[] { new OfficeChartSeries("Series", new[] { 1D }) });
        Assert.Throws<ArgumentException>(() =>
            PowerPointUtils.ValidateSharedChartData(nonFiniteScatterCategory,
                OfficeChartKind.Scatter));
    }

    [Fact]
    [Trait("Category", "MicrosoftOfficeInteroperability")]
    public void DesktopPowerPointOpensAndRendersAdvancedCorpusWhenRequested() {
        if (!string.Equals(Environment.GetEnvironmentVariable(
                "OFFICEIMO_RUN_POWERPOINT_DESKTOP_CORPUS"), "1",
                StringComparison.Ordinal)) return;

        string? configuredOutput = Environment.GetEnvironmentVariable(
            "OFFICEIMO_POWERPOINT_DESKTOP_CORPUS_OUTPUT");
        string root = string.IsNullOrWhiteSpace(configuredOutput)
            ? Path.Combine(Path.GetTempPath(),
                "OfficeIMO-PowerPointAdvancedDesktop-"
                + Guid.NewGuid().ToString("N"))
            : Path.GetFullPath(configuredOutput);
        bool deleteOutput = string.IsNullOrWhiteSpace(configuredOutput);
        Directory.CreateDirectory(root);
        try {
            PowerPointReferenceRenderResult result =
                PowerPointDesktopReferenceRenderer.TryRender(FixturePath,
                    Path.Combine(root, "source"), enabled: true);
            Assert.Equal(PowerPointReferenceRenderStatus.Succeeded,
                result.Status);
            Assert.Equal(11, result.ImagePaths.Count);

            string roundTrip = Path.Combine(root, "roundtrip.pptx");
            using (PowerPointPresentation presentation =
                   PowerPointPresentation.Load(FixturePath)) {
                ApplyRepresentativeTypedEdits(presentation);
                presentation.Save(roundTrip);
            }
            PowerPointReferenceRenderResult roundTripResult =
                PowerPointDesktopReferenceRenderer.TryRender(roundTrip,
                    Path.Combine(root, "roundtrip"), enabled: true);
            Assert.Equal(PowerPointReferenceRenderStatus.Succeeded,
                roundTripResult.Status);
            Assert.Equal(11, roundTripResult.ImagePaths.Count);
        } finally {
            if (deleteOutput) {
                try { Directory.Delete(root, recursive: true); } catch { }
            }
        }
    }

    private static void ApplyRepresentativeTypedEdits(
        PowerPointPresentation presentation) {
        PowerPointChart[] editableCharts = presentation.Slides
            .SelectMany(slide => slide.Charts)
            .Where(chart => chart.InspectImportedContent().Support ==
                PowerPointImportedChartSupport
                    .EditableWithProjectedRendering)
            .ToArray();
        Assert.Equal(6, editableCharts.Length);
        foreach (PowerPointChart editedChart in editableCharts) {
            PowerPointImportedChartReport report =
                editedChart.InspectImportedContent();
            Assert.True(editedChart.TryGetOfficeSnapshot(
                out OfficeChartSnapshot snapshot));
            OfficeChartSeries[] series = snapshot.Data.Series
                .Select((item, index) => new OfficeChartSeries(
                    index == 0 ? "OfficeIMO edited " + report.Family
                        : item.Name,
                    item.Values)).ToArray();
            editedChart.UpdateData(new OfficeChartData(
                snapshot.Data.Categories, series));
            Assert.Equal(report.Family,
                editedChart.InspectImportedContent().Family);
        }

        PowerPointSmartArt smartArt = Assert.Single(presentation.Slides
            .SelectMany(slide => slide.SmartArts));
        Assert.True(smartArt.TryGetTopology(
            out PowerPointSmartArtTopology topology,
            out string topologyDiagnostic), topologyDiagnostic);
        PowerPointSmartArtNode[] reversed = topology.Nodes.Reverse().ToArray();
        for (uint index = 0; index < reversed.Length; index++)
            reversed[index].Order = index;
        smartArt.UpdateTopology(reversed);

        PowerPointMedia audio = Assert.Single(presentation.Slides
            .SelectMany(slide => slide.Media));
        Assert.Equal(PowerPointMediaKind.Audio, audio.Kind);
        Assert.Equal(PowerPointMediaSourceKind.Embedded, audio.SourceKind);
        byte[] audioBytes = audio.GetData();
        Assert.True(audioBytes.Length > 44);
        using (var replacement = new MemoryStream(audioBytes,
                   writable: false))
            audio.UpdateData(replacement);
        audio.SetPlaybackOptions(new PowerPointMediaPlaybackOptions {
            VolumePercent = 65,
            Loop = true,
            ShowWhenStopped = true,
            TrimStartMilliseconds = 50,
            TrimEndMilliseconds = 900,
            FadeInMilliseconds = 100,
            FadeOutMilliseconds = 100
        });
    }

    private static PowerPointChart GetBar3DChart(
        PowerPointPresentation presentation) => presentation.Slides
        .SelectMany(slide => slide.Charts).Single(chart =>
            chart.InspectImportedContent().Family == "bar3DChart");

    private static ChartPart GetBar3DChartPart(
        PowerPointPresentation presentation) => presentation.OpenXmlDocument
        .PresentationPart!.SlideParts.SelectMany(slidePart =>
            slidePart.ChartParts).Single(part =>
                part.ChartSpace.Descendants<Bar3DChart>().Any());

    private static void SetWorkbookCellStyle(EmbeddedPackagePart package,
        string reference, uint styleIndex) {
        byte[] bytes = ReadPartBytes(package);
        using var stream = new MemoryStream();
        stream.Write(bytes, 0, bytes.Length);
        stream.Position = 0;
        using (SpreadsheetDocument workbook = SpreadsheetDocument.Open(
                   stream, true)) {
            WorkbookStylesPart stylesPart = workbook.WorkbookPart!
                .GetPartsOfType<WorkbookStylesPart>().Single();
            S.Stylesheet styles = stylesPart.Stylesheet!;
            S.CellFormats formats = styles.CellFormats
                ?? (styles.CellFormats = new S.CellFormats());
            while (formats.Elements<S.CellFormat>().Count() <= styleIndex) {
                formats.AppendChild(new S.CellFormat {
                    NumberFormatId = 4U,
                    FontId = 0U,
                    FillId = 0U,
                    BorderId = 0U,
                    FormatId = 0U,
                    ApplyNumberFormat = true
                });
            }
            formats.Count = (uint)formats.Elements<S.CellFormat>().Count();
            S.Cell cell = workbook.WorkbookPart
                .GetPartsOfType<WorksheetPart>().Single().Worksheet
                .Descendants<S.Cell>().Single(item =>
                    item.CellReference?.Value == reference);
            cell.StyleIndex = styleIndex;
            styles.Save();
            cell.Ancestors<S.Worksheet>().Single().Save();
        }
        stream.Position = 0;
        package.FeedData(stream);
    }

    private static uint? GetWorkbookCellStyle(EmbeddedPackagePart package,
        string reference) {
        using var stream = new MemoryStream(ReadPartBytes(package),
            writable: false);
        using SpreadsheetDocument workbook = SpreadsheetDocument.Open(
            stream, false);
        return workbook.WorkbookPart!.GetPartsOfType<WorksheetPart>()
            .Single().Worksheet.Descendants<S.Cell>().Single(item =>
                item.CellReference?.Value == reference).StyleIndex?.Value;
    }

    private static void RemoveLastWorkbookTableColumn(
        EmbeddedPackagePart package) {
        byte[] bytes = ReadPartBytes(package);
        using var stream = new MemoryStream();
        stream.Write(bytes, 0, bytes.Length);
        stream.Position = 0;
        using (SpreadsheetDocument workbook = SpreadsheetDocument.Open(
                   stream, true)) {
            TableDefinitionPart tablePart = workbook.WorkbookPart!
                .GetPartsOfType<WorksheetPart>().Single()
                .GetPartsOfType<TableDefinitionPart>().Single();
            S.TableColumns columns = tablePart.Table!.TableColumns!;
            columns.Elements<S.TableColumn>().Last().Remove();
            columns.Count = (uint)columns.Elements<S.TableColumn>().Count();
            tablePart.Table.Save();
        }
        stream.Position = 0;
        package.FeedData(stream);
    }

    private static byte[] ReadPartBytes(OpenXmlPart part) {
        using Stream stream = part.GetStream(FileMode.Open,
            FileAccess.Read);
        using var copy = new MemoryStream();
        stream.CopyTo(copy);
        return copy.ToArray();
    }

    private static string ComputePartSha256(OpenXmlPart part) {
        using Stream stream = part.GetStream(FileMode.Open,
            FileAccess.Read);
        using SHA256 sha256 = SHA256.Create();
        return string.Concat(sha256.ComputeHash(stream)
            .Select(value => value.ToString("x2")));
    }

    private static Dictionary<string, string> CaptureBar3DWorkbookShell(
        PowerPointPresentation presentation) {
        ChartPart chartPart = GetBar3DChartPart(presentation);
        EmbeddedPackagePart package = Assert.Single(
            chartPart.GetPartsOfType<EmbeddedPackagePart>());
        using var packageBytes = new MemoryStream();
        using (Stream stream = package.GetStream(FileMode.Open,
                   FileAccess.Read))
            stream.CopyTo(packageBytes);
        packageBytes.Position = 0;
        using var archive = new ZipArchive(packageBytes, ZipArchiveMode.Read,
            leaveOpen: false);
        string[] preservedEntries = {
            "xl/styles.xml",
            "xl/theme/theme1.xml",
            "xl/_rels/workbook.xml.rels"
        };
        return preservedEntries.ToDictionary(name => name, name => {
            ZipArchiveEntry entry = archive.GetEntry(name)
                ?? throw new InvalidDataException(
                    $"Microsoft chart workbook is missing {name}.");
            using Stream stream = entry.Open();
            using SHA256 sha256 = SHA256.Create();
            return string.Concat(sha256.ComputeHash(stream)
                .Select(value => value.ToString("x2")));
        }, StringComparer.Ordinal);
    }

    private static string[] GetPersistedSmartArtNodeTextsByHorizontalPosition(
        PowerPointPresentation presentation) {
        DiagramPersistLayoutPart drawingPart = Assert.Single(presentation
            .OpenXmlDocument.PresentationPart!.SlideParts
            .SelectMany(slidePart =>
                slidePart.GetPartsOfType<DiagramPersistLayoutPart>()));
        XDocument drawing;
        using (Stream stream = drawingPart.GetStream(FileMode.Open,
                   FileAccess.Read))
            drawing = XDocument.Load(stream);

        return drawing.Descendants()
            .Where(element => element.Name.LocalName == "sp")
            .Select(shape => new {
                Text = shape.Descendants().FirstOrDefault(element =>
                    element.Name.LocalName == "t")?.Value,
                X = shape.Descendants().FirstOrDefault(element =>
                    element.Name.LocalName == "off")?.Attribute("x")?.Value
            })
            .Where(shape => !string.IsNullOrWhiteSpace(shape.Text)
                            && long.TryParse(shape.X, out _))
            .OrderBy(shape => long.Parse(shape.X!))
            .Select(shape => shape.Text!)
            .ToArray();
    }

    private static void AssertMicrosoftPowerPointProducer(
        PowerPointPresentation presentation) {
        string? application = presentation.OpenXmlDocument
            .ExtendedFilePropertiesPart?.Properties?.Application?.Text;
        string? version = presentation.OpenXmlDocument
            .ExtendedFilePropertiesPart?.Properties?.ApplicationVersion?.Text;
        Assert.Equal("Microsoft Office PowerPoint", application);
        Assert.StartsWith("16.", version, StringComparison.Ordinal);
        Assert.Equal(
            "b5c2480605d376c3550941b7cc1e1601b54f881803c9655fa521211e6508e78c",
            ComputeSha256(FixturePath));
    }

    private static string ComputeSha256(string path) {
        using SHA256 sha256 = SHA256.Create();
        using FileStream stream = File.OpenRead(path);
        return string.Concat(sha256.ComputeHash(stream)
            .Select(value => value.ToString("x2")));
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
}
