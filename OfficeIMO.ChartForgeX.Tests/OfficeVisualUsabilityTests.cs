using System;
using System.IO;
using System.Linq;
using System.Text;
using global::ChartForgeX.Core;
using global::ChartForgeX.Primitives;
using global::ChartForgeX.Rendering;
using global::ChartForgeX.Themes;
using global::ChartForgeX.VisualArtifacts;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.ChartForgeX.Tests;

public sealed partial class OfficeVisualIntegrationTests {
    [Fact]
    public void ExplicitBoundsContainByDefaultAndStretchOnlyWhenRequested() {
        var source = new OfficeVisualSource("<svg xmlns='http://www.w3.org/2000/svg' width='200' height='100'><rect width='200' height='100' fill='#123456'/></svg>");
        var contained = source.ToOfficeVisual(new OfficeVisualConversionOptions { WidthPoints = 100D, HeightPoints = 100D });
        var stretched = source.ToOfficeVisual(new OfficeVisualConversionOptions { WidthPoints = 100D, HeightPoints = 100D, Fit = OfficeImageFit.Stretch });
        Assert.Equal(100D, contained.WidthPoints);
        Assert.Equal(50D, contained.HeightPoints);
        Assert.Equal(100D, stretched.HeightPoints);
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeVisualConversionOptions { Fit = OfficeImageFit.Cover });
    }

    [Fact]
    public void ResizingReusesPayloadAndFidelityWhileScalingRegionsWithoutMutatingSource() {
        var original = CreateArtifact().ToOfficeVisual();
        var resized = original.WithSize(original.WidthPoints / 2D);
        Assert.Equal(original.GetPlacementBytes(), resized.GetPlacementBytes());
        Assert.Equal(original.GetSvgBytes(), resized.GetSvgBytes());
        Assert.Same(original.Report, resized.Report);
        Assert.Equal(original.AlternativeText, resized.AlternativeText);
        Assert.Equal(original.HeightPoints / 2D, resized.Drawing.Height);
        Assert.Equal(original.WidthPoints, original.Drawing.Width);
        var originalRegion = Assert.Single(original.Regions);
        var resizedRegion = Assert.Single(resized.Regions);
        Assert.Equal(originalRegion.Left / 2D, resizedRegion.Left);
        Assert.Equal(originalRegion.Top / 2D, resizedRegion.Top);
        Assert.Equal(originalRegion.Width / 2D, resizedRegion.Width);
        Assert.Equal(originalRegion.Href, resizedRegion.Href);
    }

    [Fact]
    public void DocumentStylePreparesAtPlacementScaleWithSharedThemeAndReadableLabels() {
        var style = new OfficeVisualDocumentStyle("Arial", labelSizePoints: 12D);
        var context = style.CreateContext(300D, 180D, VisualThemeMode.Dark);
        Assert.Equal(400D, context.Layout.Size.Width);
        Assert.Equal(240D, context.Layout.Size.Height);
        Assert.Equal(12D, context.Theme.Typography.AxisSize * .75D);
        Assert.Equal(style.HeadingSizePoints, context.Theme.Typography.TitleSize * .75D);
        Assert.Equal(style.FontFamily, context.Theme.Typography.Family);
        Assert.Equal(VisualTheme.Graphite().Resolve(VisualThemeMode.Dark).Background,
            context.Theme.Resolve(VisualThemeMode.Dark).Background);
    }

    [Fact]
    public void SimpleInsertionsFitSavedDestinationsAndRetainAccessibility() {
        var chart = Chart.Create().AddLine("Requests", new[] { new ChartPoint(1, 10), new ChartPoint(2, 20) });
        var artifact = chart.Prepare(OfficeVisualDocumentStyle.Default.CreateContext(600D, 300D,
            frame: new VisualFrame("Requests", showLegend: false))).ToArtifact("requests");
        artifact.Accessibility.WithTextAlternative("Requests", "Two measured requests.");
        string folder = Path.Combine(Path.GetTempPath(), "OfficeIMO-VisualFit-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        try {
            string wordPath = Path.Combine(folder, "fit.docx");
            double available;
            using (var document = WordDocument.Create(wordPath)) {
                var paragraph = document.AddParagraph();
                available = paragraph.GetContentWidthPoints();
                paragraph.AddVisualArtifact(artifact);
                document.Save();
            }
            using (var document = WordDocument.Load(wordPath)) {
                var image = Assert.Single(document.Images);
                Assert.Equal(Math.Min(600D, available), image.Width!.Value * .75D, 3);
                Assert.Equal("Two measured requests.", image.Description);
                Assert.Equal(2D, image.Width.Value / image.Height!.Value, 3);
            }
            string excelPath = Path.Combine(folder, "fit.xlsx");
            int rangeWidth, rangeHeight;
            using (var workbook = ExcelDocument.Create(excelPath)) {
                var sheet = workbook.AddWorksheet("Requests");
                var size = sheet.GetRangeSizePixels("B2:D8");
                rangeWidth = size.WidthPixels; rangeHeight = size.HeightPixels;
                sheet.AddVisualArtifact("B2:D8", artifact);
                workbook.Save();
            }
            using (var workbook = ExcelDocument.Load(excelPath)) {
                var image = Assert.Single(workbook.Sheets[0].Images);
                Assert.True(image.WidthPixels <= rangeWidth && image.HeightPixels <= rangeHeight);
                Assert.InRange((double)image.WidthPixels / image.HeightPixels, 1.98D, 2.02D);
                Assert.Equal("Two measured requests.", image.Description);
            }
            string slidesPath = Path.Combine(folder, "fit.pptx");
            using (var presentation = PowerPointPresentation.Create(slidesPath)) {
                presentation.AddSlide().AddVisualArtifact(artifact, PowerPointLayoutBox.FromPoints(36D, 60D, 300D, 100D));
                presentation.Save();
            }
            using (var presentation = PowerPointPresentation.Load(slidesPath)) {
                var picture = Assert.Single(presentation.Slides[0].Pictures);
                Assert.Equal(200D, picture.WidthPoints, 3);
                Assert.Equal(100D, picture.HeightPoints, 3);
                Assert.Equal(86D, picture.LeftPoints, 3);
                Assert.Equal("Two measured requests.", picture.Description);
            }
            var pdfOptions = new PdfOptions { CompressContentStreams = false };
            pdfOptions.EnableTaggedPdfCatalogMarkers();
            byte[] pdf = PdfDocument.Create(document => document.Content(content => content.AddVisualArtifact(artifact)), pdfOptions)
                .ToBytes();
            Assert.Equal("%PDF", Encoding.ASCII.GetString(pdf, 0, 4));
            string textAlternative = BitConverter.ToString(Encoding.UTF8.GetBytes("Two measured requests.")).Replace("-", string.Empty);
            Assert.Contains("/Figure << /Alt <" + textAlternative + ">", Encoding.ASCII.GetString(pdf));
        } finally {
            Directory.Delete(folder, recursive: true);
        }
    }
}
