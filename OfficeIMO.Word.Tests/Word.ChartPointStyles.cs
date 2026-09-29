using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public sealed class WordChartPointStylesTests {
    [Fact]
    public void PointStyles_RejectConflictingDirectAndMarkerShapeOwners() {
        using WordDocument document = WordDocument.Create();
        WordChart chart = document.AddChart(OfficeChartKind.Line,
            new OfficeChartData(new[] { "A", "B" }, new[] {
                new OfficeChartSeries("Measured", new[] { 2d, 3d }) }));
        C.LineChartSeries native = chart.ChartPart!.ChartSpace!.Descendants<C.LineChartSeries>().Single();
        var point = new C.DataPoint(new C.Index { Val = 0 });
        point.AddChild(new C.Marker(new C.ChartShapeProperties(
            new A.SolidFill(new A.RgbColorModelHex { Val = "0000FF" }))), true);
        point.AddChild(new C.ChartShapeProperties(
            new A.SolidFill(new A.RgbColorModelHex { Val = "FF0000" })), true);
        native.AddChild(point, true);
        Assert.False(chart.TryGetSnapshot(out _));
    }

    [Fact]
    public void PointStyles_RejectUnprojectedNativeShapeAttributes() {
        using WordDocument document = WordDocument.Create();
        WordChart chart = document.AddChart("Status", false, 360, 180);
        chart.AddPie("Pass", 3).AddPie("Fail", 1);
        chart.SetDataPointStyle(0, 0, new(OfficeColor.White));
        C.ChartShapeProperties properties = chart.ChartPart!.ChartSpace!
            .Descendants<C.DataPoint>().Single().ChartShapeProperties!;
        properties.SetAttribute(new DocumentFormat.OpenXml.OpenXmlAttribute("", "bwMode", "", "black"));
        Assert.False(chart.TryGetSnapshot(out _));
    }

    [Fact]
    public void PointStyles_PreserveNativePieAppearanceAndPdfAfterReopen() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".docx");
        try {
            using (WordDocument authored = WordDocument.Create(path)) {
                WordChart authoredChart = authored.AddChart("Status", false, 360, 180);
                authoredChart.AddPie("Pass", 3).AddPie("Unknown", 2).AddPie("Fail", 1);
                authoredChart.SetDataPointStyle(0, 0, new(showOutline: true));
                authoredChart.SetDataPointStyle(0, 1, new(noFill: true, outlineColor: OfficeColor.Black, outlineWidth: 2));
                authoredChart.SetDataPointStyle(0, 2, new(OfficeColor.White, hatch: OfficeChartHatchPattern.DiagonalCross,
                    hatchColor: OfficeColor.Parse("#7300A3"), outlineColor: OfficeColor.Black));
                authored.Save();
            }
            using WordDocument reopened = WordDocument.Load(path);
            WordChart chart = Assert.Single(reopened.Charts);
            Assert.True(chart.TryGetSnapshot(out WordChartSnapshot snapshot));
            Assert.True(snapshot.Data.Series[0].PointStyles![0]!.ShowOutline);
            Assert.True(snapshot.Data.Series[0].PointStyles![1]!.NoFill);
            Assert.Equal(OfficeChartHatchPattern.DiagonalCross, snapshot.Data.Series[0].PointStyles![2]!.Hatch);
            Assert.Empty(reopened.ValidateDocument());
            var options = new WordToPdfOptions { IncludePageNumbers = false };
            byte[] pdf = reopened.ToPdfBytes(options);
            Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeBodyChartUnsupported");
            var page = Assert.Single(PdfCore.PdfDocument.Load(pdf).Render.Pages(options:
                new PdfCore.PdfPageRenderOptions { Dpi = 72, MaxPages = 1, ContinueOnError = false }));
            Assert.True(OfficePngReader.TryDecode(page.Bytes!, out OfficeRasterImage? raster));
            int pixels = 0;
            for (int y = 0; y < raster!.Height; y++)
                for (int x = 0; x < raster.Width; x++)
                    if (IsPurpleStroke(raster.GetPixel(x, y))) pixels++;
            Assert.True(pixels > 25, "Expected hatch strokes in the PDF; actual pixels " + pixels);
            chart.SetDataPointStyle(0, 0, null).SetDataPointStyle(0, 1, null).SetDataPointStyle(0, 2, null);
            Assert.True(chart.TryGetSnapshot(out snapshot));
            Assert.Null(snapshot.Data.Series[0].PointStyles);
        } finally { if (File.Exists(path)) File.Delete(path); }
    }
    // Thin hatch strokes blend with their background during antialiasing.
    private static bool IsPurpleStroke(OfficeColor pixel) =>
        pixel.R < 200 && pixel.G < 150 && pixel.B > pixel.G + 40 && pixel.R > pixel.G + 20;
}
