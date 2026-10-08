using System;
using System.IO;
using System.Text;
using ChartForgeX.Core;
using ChartForgeX.Primitives;
using ChartForgeX.Rendering;
using ChartForgeX.VisualArtifacts;
using OfficeIMO.Excel;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.ChartForgeX.Tests;

public sealed partial class OfficeVisualIntegrationTests {
    [Fact]
    public void PreparedChartPlacementRetainsDimensionsAndAccessibleTextThroughSavedDocuments() {
        const string description = "Service load increases across three samples.";
        var chart = Chart.Create().AddLine("Load", new[] { new ChartPoint(0, 20), new ChartPoint(1, 35), new ChartPoint(2, 50) })
            .WithAccessibility(value => value.WithTextAlternative("Service load", description, "en"));
        var prepared = chart.Prepare(new VisualRenderContext(
            layout: new VisualLayoutOptions(new VisualSize(320, 200), padding: 16),
            frame: new VisualFrame(title: "Service load", showLegend: false)));
        var artifact = prepared.ToArtifact("service-load", VisualArtifactKind.Chart);
        var conversion = artifact.ToOfficeVisual(new OfficeVisualConversionOptions { PointsPerPixel = 0.5D });
        Assert.Equal(160D, conversion.WidthPoints, 6);
        Assert.Equal(100D, conversion.HeightPoints, 6);
        Assert.Equal(description, conversion.AlternativeText);

        string folder = Path.Combine(Path.GetTempPath(), "OfficeIMO-ChartForgeX-prepared-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        string wordPath = Path.Combine(folder, "prepared.docx");
        string excelPath = Path.Combine(folder, "prepared.xlsx");
        string presentationPath = Path.Combine(folder, "prepared.pptx");
        string pdfPath = Path.Combine(folder, "prepared.pdf");
        try {
            using (var document = WordDocument.Create(wordPath)) {
                document.AddParagraph().AddVisualArtifact(conversion);
                document.Save();
            }
            using (var document = ExcelDocument.Create(excelPath)) {
                document.AddWorksheet("Load").AddVisualArtifact(2, 3, conversion);
                document.Save();
            }
            using (var document = PowerPointPresentation.Create(presentationPath)) {
                document.AddSlide().AddVisualArtifact(conversion, 24D, 36D);
                document.Save();
            }
            PdfDocument.Create(_ => { }, new PdfOptions { CompressContentStreams = false }.EnableTaggedPdfCatalogMarkers())
                .AddVisualArtifact(conversion).Save(pdfPath);

            using (var document = WordDocument.Load(wordPath)) {
                var image = Assert.Single(document.Images);
                Assert.Equal(160D / 0.75D, image.Width!.Value, 6);
                Assert.Equal(100D / 0.75D, image.Height!.Value, 6);
                Assert.Equal(description, image.Description);
                Assert.Equal("Service load", image.Title);
            }
            using (var document = ExcelDocument.Load(excelPath)) {
                var image = Assert.Single(document.Sheets[0].Images);
                Assert.Equal(213, image.WidthPixels);
                Assert.Equal(133, image.HeightPixels);
                Assert.Equal(description, image.Description);
                Assert.Equal("Service load", image.Title);
                Assert.Equal(2, image.RowIndex);
                Assert.Equal(3, image.ColumnIndex);
            }
            using (var document = PowerPointPresentation.Load(presentationPath)) {
                var picture = Assert.Single(document.Slides[0].Pictures);
                Assert.Equal(160D, picture.WidthPoints, 6);
                Assert.Equal(100D, picture.HeightPoints, 6);
                Assert.Equal(description, picture.Description);
                Assert.Equal("Service load", picture.Title);
            }
            var pdf = File.ReadAllBytes(pdfPath);
            Assert.Single(PdfReadDocument.Open(pdf).Pages);
            string textAlternative = BitConverter.ToString(Encoding.UTF8.GetBytes(description)).Replace("-", string.Empty);
            Assert.Contains("/Figure << /Alt <" + textAlternative + ">", Encoding.ASCII.GetString(pdf), StringComparison.Ordinal);
        } finally { if (Directory.Exists(folder)) Directory.Delete(folder, recursive: true); }
    }

    private static uint PngPixelsPerMeter(byte[] png) {
        uint ReadUInt32(int offset) => ((uint)png[offset] << 24) | ((uint)png[offset + 1] << 16) | ((uint)png[offset + 2] << 8) | png[offset + 3];
        for (int offset = 8; offset + 12 <= png.Length; offset += checked((int)ReadUInt32(offset)) + 12) {
            if (Encoding.ASCII.GetString(png, offset + 4, 4) != "pHYs") continue;
            Assert.Equal(9U, ReadUInt32(offset));
            Assert.Equal(1, png[offset + 16]);
            uint density = ReadUInt32(offset + 8);
            Assert.Equal(density, ReadUInt32(offset + 12));
            return density;
        }
        throw new InvalidOperationException("The PNG does not declare pixel density.");
    }
}
