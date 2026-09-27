using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelChartPointStylesTests {
    [Theory]
    [InlineData(OfficeChartKind.Pie)]
    [InlineData(OfficeChartKind.Doughnut)]
    [InlineData(OfficeChartKind.ColumnClustered)]
    public void PointStyles_PreserveNativeStylesAndRenderHatchesAfterReopen(OfficeChartKind kind) {
        var hatch = new OfficeChartPointStyle(OfficeColor.White, hatch: OfficeChartHatchPattern.DiagonalCross,
            hatchColor: OfficeColor.Parse("#7300A3"), outlineColor: OfficeColor.Black, outlineWidth: 2);
        var series = new OfficeChartSeries("Status", new[] { 3d, 2d, 1d })
            .WithPointStyles(new OfficeChartPointStyle?[] { new(showOutline: true), new(noFill: true, outlineColor: OfficeColor.Black), hatch });
        using ExcelDocument authored = ExcelDocument.Create();
        authored.AddWorksheet("Results").AddChart(kind,
            new OfficeChartData(new[] { "Pass", "Unknown", "Fail" }, new[] { series }), 1, 1);
        using var bytes = new MemoryStream(authored.ToBytes());
        using ExcelDocument reopened = ExcelDocument.Load(bytes);
        Assert.Empty(reopened.ValidateDocument());
        ExcelChart chart = Assert.Single(reopened.Sheets.Single(sheet => sheet.Name == "Results").Charts);
        Assert.True(chart.TryGetSnapshot(out ExcelChartSnapshot snapshot));
        Assert.True(snapshot.Data.Series[0].PointStyles![0]!.ShowOutline);
        Assert.True(snapshot.Data.Series[0].PointStyles![1]!.NoFill);
        Assert.Equal(hatch.Hatch, snapshot.Data.Series[0].PointStyles![2]!.Hatch);
        string svg = System.Text.Encoding.UTF8.GetString(chart.ExportImage(OfficeImageExportFormat.Svg).Bytes);
        Assert.Contains("#7300A3", svg, System.StringComparison.OrdinalIgnoreCase);
        Assert.Contains("clipPath", svg, System.StringComparison.Ordinal);
        byte[] pdf = reopened.ToPdfBytes(new ExcelToPdfOptions { IncludeSheetHeadings = false });
        var pages = OfficeIMO.Pdf.PdfDocument.Load(pdf).Render.Pages(options:
            new OfficeIMO.Pdf.PdfPageRenderOptions { Dpi = 72, MaxPages = 8, ContinueOnError = false });
        int pixels = 0;
        foreach (var page in pages) {
            Assert.True(OfficePngReader.TryDecode(page.Bytes!, out OfficeRasterImage? raster));
            for (int y = 0; y < raster!.Height; y++)
                for (int x = 0; x < raster.Width; x++)
                    if (IsPurpleStroke(raster.GetPixel(x, y))) pixels++;
        }
        Assert.True(pixels > 25, "Expected hatch strokes in the workbook PDF; actual pixels " + pixels);
        chart.SetDataPointStyle(0, 0, null).SetDataPointStyle(0, 1, null).SetDataPointStyle(0, 2, null);
        Assert.True(chart.TryGetSnapshot(out snapshot));
        Assert.Null(snapshot.Data.Series[0].PointStyles);
    }
    // Thin hatch strokes blend with their background during antialiasing.
    private static bool IsPurpleStroke(OfficeColor pixel) =>
        pixel.R < 200 && pixel.G < 150 && pixel.B > pixel.G + 40 && pixel.R > pixel.G + 20;
}
