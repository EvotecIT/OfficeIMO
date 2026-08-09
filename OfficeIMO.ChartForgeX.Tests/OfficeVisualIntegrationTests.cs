using System;
using System.IO;
using global::ChartForgeX.Primitives;
using global::ChartForgeX.VisualArtifacts;
using OfficeIMO.ChartForgeX;
using OfficeIMO.Excel;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.ChartForgeX.Tests;

public sealed class OfficeVisualIntegrationTests {
    [Fact]
    public void ConversionPreservesVectorPayloadDimensionsAccessibilityAndRegions() {
        VisualArtifact artifact = CreateArtifact();
        var options = new OfficeVisualConversionOptions { WidthPoints = 360D };

        OfficeVisualConversionResult result = artifact.ToOfficeVisual(options);

        Assert.True(result.GetSvgBytes().Length > 100);
        Assert.True(result.GetPlacementBytes().Length > 100);
        Assert.True(result.WidthPoints == 360D);
        Assert.True(result.HeightPoints > 0D);
        Assert.Equal("Quarterly service health across API and worker tiers.", result.AlternativeText);
        Assert.True(result.Drawing.Width == result.WidthPoints);
        Assert.True(result.Drawing.Height == result.HeightPoints);
        Assert.Single(result.Regions);
        Assert.Equal("https://example.test/api", result.Regions[0].Href);
        Assert.NotNull(result.Regions[0].Width);
        Assert.True(result.Report.IsVector || result.Report.UnsupportedSvgFeatureCount > 0);
        Assert.Equal(OfficeVisualMediaFormat.Svg, result.PlacementFormat);
    }

    [Fact]
    public void RasterizeWhenNeededSelectsThePlacementPayloadReportedByConversion() {
        OfficeVisualConversionResult result = CreateArtifact().ToOfficeVisual(new OfficeVisualConversionOptions {
            SvgPolicy = OfficeVisualSvgPolicy.RasterizeWhenNeeded
        });

        if (result.Report.UsedRasterFallback) {
            Assert.Equal(OfficeVisualMediaFormat.Png, result.PlacementFormat);
            byte[] payload = result.GetPlacementBytes();
            Assert.Equal(0x89, payload[0]);
            Assert.Equal((byte)'P', payload[1]);
        } else {
            Assert.Equal(OfficeVisualMediaFormat.Svg, result.PlacementFormat);
            Assert.Contains("<svg", System.Text.Encoding.UTF8.GetString(result.GetPlacementBytes()));
        }
    }

    [Fact]
    public void PortableSvgSourceSupportsCrossProcessAndPowerShellExchange() {
        byte[] svg = CreateArtifact().ToOfficeVisual().GetSvgBytes();
        var source = new OfficeVisualSource(svg) {
            Id = "portable-health",
            Title = "Portable Health",
            AlternativeText = "Portable service health visual."
        };

        OfficeVisualConversionResult result = source.ToOfficeVisual(new OfficeVisualConversionOptions {
            WidthPoints = 280D,
            SvgPolicy = OfficeVisualSvgPolicy.RasterizeWhenNeeded
        });

        Assert.Null(result.Artifact);
        Assert.Equal("portable-health", result.Id);
        Assert.Equal("Portable Health", result.Title);
        Assert.Equal("Portable service health visual.", result.AlternativeText);
        Assert.Equal(280D, result.WidthPoints);
        Assert.True(result.GetPlacementBytes().Length > 100);

        var titledSource = new OfficeVisualSource(svg) { Id = "fallback-id", Title = "Accessible fallback title" };
        OfficeVisualConversionResult titledResult = titledSource.ToOfficeVisual();
        Assert.Equal("Accessible fallback title", titledResult.AlternativeText);

        var markupSource = new OfficeVisualSource(System.Text.Encoding.UTF8.GetString(svg));
        Assert.Equal(svg, markupSource.GetSvgBytes());
    }

    [Fact]
    public void PlacementCreatesReadableWordExcelPowerPointAndPdfPackages() {
        string folder = Path.Combine(Path.GetTempPath(), "OfficeIMO-ChartForgeX-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        try {
            var renderOptions = new VisualArtifactRenderOptions();
            renderOptions.Watermarks.Add(VisualWatermark.FromText("INTERNAL"));
            VisualWatermark imageWatermark = VisualWatermark.FromImage(
                CreateArtifact().ToPng(),
                "image/png");
            imageWatermark.Width = 18D;
            imageWatermark.Height = 18D;
            imageWatermark.Padding = 8D;
            renderOptions.Watermarks.Add(imageWatermark);
            OfficeVisualConversionResult visual = CreateArtifact().ToOfficeVisual(
                new OfficeVisualConversionOptions {
                    WidthPoints = 300D,
                    RenderOptions = renderOptions
                });
            string wordPath = Path.Combine(folder, "visual.docx");
            string excelPath = Path.Combine(folder, "visual.xlsx");
            string powerPointPath = Path.Combine(folder, "visual.pptx");
            string pdfPath = Path.Combine(folder, "visual.pdf");

            using (WordDocument document = WordDocument.Create(wordPath)) {
                document.AddParagraph().AddVisualArtifact(visual);
                document.Save();
            }

            using (ExcelDocument workbook = ExcelDocument.Create(excelPath)) {
                ExcelSheet sheet = workbook.AddWorksheet("Dashboard");
                sheet.AddVisualArtifact(2, 2, visual);
                workbook.Save();
            }

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(powerPointPath)) {
                presentation.AddSlide().AddVisualArtifact(visual, 36D, 54D);
                presentation.Save();
            }

            PdfDocument pdfDocument = PdfDocument.Create(_ => { });
            pdfDocument.AddVisualArtifact(visual).Save(pdfPath);

            using (WordDocument document = WordDocument.Load(wordPath)) {
                Assert.Single(document.Images);
            }
            using (ExcelDocument workbook = ExcelDocument.Load(excelPath)) {
                Assert.Single(workbook.Sheets[0].Images);
            }
            using (PowerPointPresentation presentation = PowerPointPresentation.Load(powerPointPath)) {
                Assert.Single(presentation.Slides[0].Pictures);
            }
            byte[] pdf = File.ReadAllBytes(pdfPath);
            Assert.True(pdf.Length > 100);
            Assert.Equal("%PDF", System.Text.Encoding.ASCII.GetString(pdf, 0, 4));
            Assert.Single(PdfReadDocument.Open(pdf).Pages);
        } finally {
            if (Directory.Exists(folder)) Directory.Delete(folder, recursive: true);
        }
    }

    [Fact]
    public void ConversionOptionsRejectInvalidSizingAndSvgLimits() {
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeVisualConversionOptions { PointsPerPixel = 0D });
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeVisualConversionOptions { WidthPoints = double.NaN });
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeVisualConversionOptions { MaximumSvgElements = 0 });
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeVisualConversionOptions { SvgPolicy = (OfficeVisualSvgPolicy)999 });
    }

    private static VisualArtifact CreateArtifact() {
        var table = TableArtifact.Create("service-health")
            .WithTitle("Service Health")
            .WithSubtitle("Current production status")
            .AddColumn("service", "Service")
            .AddColumn("state", "State", TableArtifactColumnType.Status)
            .AddRow("api", "API", "Healthy")
            .AddRow("worker", "Worker", "Warning");
        VisualArtifact artifact = table.ToVisualArtifact();
        artifact.Accessibility.WithTextAlternative(
            "Service health",
            "Quarterly service health across API and worker tiers.",
            "en");
        artifact.Regions.Add(new VisualArtifactRegion {
            Id = "api",
            Kind = "row",
            Label = "API",
            AlternativeText = "API service details",
            Href = "https://example.test/api",
            Bounds = new ChartRect(10D, 40D, 160D, 24D)
        });
        return artifact;
    }
}
