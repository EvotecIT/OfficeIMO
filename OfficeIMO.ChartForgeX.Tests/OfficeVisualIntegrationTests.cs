using System;
using System.IO;
using System.Linq;
using System.Text;
using global::ChartForgeX.Primitives;
using global::ChartForgeX.VisualArtifacts;
using OfficeIMO.ChartForgeX;
using OfficeIMO.Drawing;
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

        const string rectangularViewport = "<svg xmlns='http://www.w3.org/2000/svg' width='200' height='100' viewBox='0 0 100 100'><rect width='100' height='100' fill='#2563eb'/></svg>";
        OfficeVisualConversionResult rectangular = new OfficeVisualSource(rectangularViewport).ToOfficeVisual();
        Assert.Equal(150D, rectangular.WidthPoints, 6);
        Assert.Equal(75D, rectangular.HeightPoints, 6);
        Assert.Equal(150D, rectangular.Drawing.Width, 6);
        Assert.Equal(75D, rectangular.Drawing.Height, 6);
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
    public void PdfPlacementCarriesGeneratedAlternativeTextWithoutMutatingCallerStyle() {
        OfficeVisualConversionResult visual = CreateArtifact().ToOfficeVisual(new OfficeVisualConversionOptions {
            WidthPoints = 300D
        });
        var style = new PdfDrawingStyle { SpacingBefore = 6D };

        byte[] pdf = PdfDocument.Create(
                _ => { },
                new PdfOptions { CompressContentStreams = false }.EnableTaggedPdfCatalogMarkers())
            .AddVisualArtifact(visual, style: style)
            .ToBytes();
        string content = Encoding.ASCII.GetString(pdf);

        Assert.Null(style.AlternativeText);
        Assert.Contains(
            "/Figure << /Alt <517561727465726C792073657276696365206865616C7468206163726F73732041504920616E6420776F726B65722074696572732E>",
            content,
            StringComparison.Ordinal);
    }

    [Fact]
    public void DecorativeArtifactUsesPdfArtifactStructureByDefault() {
        VisualArtifact artifact = CreateArtifact();
        artifact.Accessibility.AsDecorative();
        OfficeVisualConversionResult visual = artifact.ToOfficeVisual(new OfficeVisualConversionOptions { WidthPoints = 300D });

        byte[] pdf = PdfDocument.Create(
                _ => { },
                new PdfOptions { CompressContentStreams = false }.EnableTaggedPdfCatalogMarkers())
            .AddVisualArtifact(visual, style: new PdfDrawingStyle { SpacingBefore = 6D })
            .ToBytes();
        string content = Encoding.ASCII.GetString(pdf);

        Assert.True(visual.IsDecorative);
        Assert.Empty(visual.AlternativeText);
        Assert.Contains("/Artifact", content, StringComparison.Ordinal);
        Assert.DoesNotContain("/Figure << /Alt", content, StringComparison.Ordinal);
    }

    [Fact]
    public void DecorativeArtifactMarksPowerPointPictureWithoutInformativeMetadata() {
        VisualArtifact artifact = CreateArtifact();
        artifact.Accessibility.AsDecorative();
        OfficeVisualConversionResult visual = artifact.ToOfficeVisual(new OfficeVisualConversionOptions { WidthPoints = 300D });
        string path = Path.Combine(Path.GetTempPath(), "OfficeIMO-ChartForgeX-decorative-" + Guid.NewGuid().ToString("N") + ".pptx");
        string wordPath = Path.ChangeExtension(path, ".docx");
        try {
            using var presentation = PowerPointPresentation.Create(path);
            PowerPointPicture picture = presentation.AddSlide().AddVisualArtifact(visual, 36D, 54D);

            Assert.True(picture.Decorative);
            Assert.Null(picture.Title);
            Assert.Null(picture.Description);

            using var document = WordDocument.Create(wordPath);
            WordImage image = document.AddParagraph().AddVisualArtifact(visual);
            Assert.Null(image.Title);
            Assert.True(string.IsNullOrEmpty(image.Description));
        } finally {
            if (File.Exists(path)) File.Delete(path);
            if (File.Exists(wordPath)) File.Delete(wordPath);
        }
    }

    [Fact]
    public void PdfPlacementRasterizesUnsupportedEffectGroupsOrFailsClosed() {
        const string blendedSvg = "<svg xmlns='http://www.w3.org/2000/svg' width='120' height='80' viewBox='0 0 120 80'><rect width='120' height='80' fill='#ffffff'/><g style='mix-blend-mode:multiply'><rect x='10' y='10' width='70' height='50' fill='#ef4444'/><rect x='40' y='20' width='70' height='50' fill='#2563eb'/></g></svg>";
        const string maskedSvg = "<svg xmlns='http://www.w3.org/2000/svg' width='120' height='80' viewBox='0 0 120 80'><defs><mask id='fade' maskUnits='userSpaceOnUse' x='0' y='0' width='120' height='80'><rect width='60' height='80' fill='white'/></mask></defs><rect width='120' height='80' fill='#2563eb' mask='url(#fade)'/></svg>";
        foreach (string svg in new[] { blendedSvg, maskedSvg }) {
            var source = new OfficeVisualSource(svg) { AlternativeText = "Effect-group visual" };
            OfficeVisualConversionResult preserved = source.ToOfficeVisual(new OfficeVisualConversionOptions { SvgPolicy = OfficeVisualSvgPolicy.PreserveVector });
            Assert.Contains(preserved.Drawing.Elements, element => element is OfficeDrawingEffectGroup);
            OfficeRasterImage expected = OfficeDrawingRasterRenderer.Render(preserved.Drawing, scale: 96D / 72D);

            byte[] pdf = PdfDocument.Create(_ => { }, new PdfOptions { CompressContentStreams = false })
                .AddVisualArtifact(preserved)
                .ToBytes();
            Assert.Equal("%PDF", Encoding.ASCII.GetString(pdf, 0, 4));
            Assert.Contains("/Subtype /Image", Encoding.ASCII.GetString(pdf), StringComparison.Ordinal);
            PdfExtractedImage embedded = Assert.Single(PdfDocument.Open(pdf).Read.Images());
            Assert.True(OfficeRasterImageDecoder.TryDecode(embedded.Bytes, out OfficeRasterImage? actual));
            Assert.NotNull(actual);
            Assert.Equal(expected.Width, actual!.Width);
            Assert.Equal(expected.Height, actual.Height);
            Assert.True(expected.GetPixels().SequenceEqual(actual.GetPixels()), "PDF fallback pixels should come from the effect-capable OfficeDrawing raster renderer.");
        }

        var requiredSource = new OfficeVisualSource(blendedSvg);
        OfficeVisualConversionResult required = requiredSource.ToOfficeVisual(new OfficeVisualConversionOptions { SvgPolicy = OfficeVisualSvgPolicy.RequireVector });
        Assert.Throws<NotSupportedException>(() => PdfDocument.Create(_ => { }).AddVisualArtifact(required));
    }

    [Fact]
    public void SymbolUseGroupsRemainVectorAndCanBeWrittenToPdf() {
        const string svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"120\" height=\"80\" viewBox=\"0 0 120 80\"><defs><symbol id=\"badge\" viewBox=\"0 0 40 30\"><rect width=\"40\" height=\"30\" rx=\"4\" fill=\"#2563eb\"/><text x=\"20\" y=\"19\" text-anchor=\"middle\" fill=\"white\">OK</text></symbol></defs><use href=\"#badge\" x=\"20\" y=\"15\" width=\"80\" height=\"50\"/></svg>";
        var source = new OfficeVisualSource(svg) {
            AlternativeText = "Reusable status badge"
        };

        OfficeVisualConversionResult visual = source.ToOfficeVisual(new OfficeVisualConversionOptions {
            SvgPolicy = OfficeVisualSvgPolicy.RequireVector
        });

        Assert.Contains(
            visual.Drawing.Elements,
            element => element is OfficeDrawingGroup || element is OfficeDrawingEffectGroup);
        byte[] pdf = PdfDocument.Create(_ => { }, new PdfOptions { CompressContentStreams = false })
            .AddVisualArtifact(visual)
            .ToBytes();
        Assert.Equal("%PDF", Encoding.ASCII.GetString(pdf, 0, 4));
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
