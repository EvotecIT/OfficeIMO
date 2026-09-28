using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class ExcelChartPointStylesTests {
    [Fact]
    public void MarkerOwnedPointHatchRemainsExportable() {
        var data = new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeChartSeries("Line", new[] { 1d, 2d })
        });
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Markers");
        ExcelChart chart = sheet.AddChart(OfficeChartKind.Line, data, 1, 4);
        ChartPart part = Assert.Single(sheet.WorksheetPart.DrawingsPart!.ChartParts);
        C.LineChartSeries nativeSeries = Assert.Single(part.ChartSpace!.Descendants<C.LineChartSeries>());
        nativeSeries.InsertBefore(new C.DataPoint(new C.Index { Val = 0U },
            new C.Marker(new C.ChartShapeProperties())), nativeSeries.GetFirstChild<C.CategoryAxisData>());
        var hatch = new OfficeChartPointStyle(OfficeColor.White,
            hatch: OfficeChartHatchPattern.Cross, hatchColor: OfficeColor.Parse("#7300A3"));
        chart.SetDataPointStyle(0, 0, hatch);

        Assert.NotNull(Assert.Single(nativeSeries.Elements<C.DataPoint>())
            .GetFirstChild<C.Marker>()!.ChartShapeProperties!.GetFirstChild<A.PatternFill>());
        Assert.True(chart.TryGetSnapshot(out ExcelChartSnapshot snapshot));
        Assert.Equal(hatch.Hatch, snapshot.Data.Series[0].PointStyles![0]!.Hatch);
        Assert.Contains("#7300A3", System.Text.Encoding.UTF8.GetString(
            chart.ExportImage(OfficeImageExportFormat.Svg).Bytes), System.StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void SharedComboPointStylesFollowNativeChartIndexesAcrossLayers() {
        var data = new OfficeChartData(new[] { "Q1", "Q2" }, new[] {
            new OfficeChartSeries("Columns A", new[] { 12D, 18D }, null, null, null, true,
                renderKind: OfficeChartKind.ColumnClustered),
            new OfficeChartSeries("Trend", new[] { 14D, 20D }, null, null, null, true,
                renderKind: OfficeChartKind.Line)
                .WithPointStyles(new OfficeChartPointStyle?[] {
                    new(fillColor: OfficeColor.Parse("#AABBCC")), null
                }),
            new OfficeChartSeries("Columns B", new[] { 10D, 16D }, null, null, null, true,
                renderKind: OfficeChartKind.ColumnClustered)
        });
        using ExcelDocument document = ExcelDocument.Create();
        document.AddWorksheet("Shared").AddChart(OfficeChartKind.ColumnClustered, data, row: 1, column: 4);

        using SpreadsheetDocument package = SpreadsheetDocument.Open(new MemoryStream(document.ToBytes()), false);
        ChartPart part = Assert.Single(package.WorkbookPart!.WorksheetParts
            .SelectMany(sheet => sheet.DrawingsPart?.ChartParts ?? Enumerable.Empty<ChartPart>()));
        C.LineChartSeries line = Assert.Single(part.ChartSpace!.Descendants<C.LineChartSeries>());
        Assert.Equal(1U, line.Index!.Val!.Value);
        A.RgbColorModelHex color = Assert.Single(line.Elements<C.DataPoint>())
            .GetFirstChild<C.ChartShapeProperties>()!.GetFirstChild<A.SolidFill>()!
            .GetFirstChild<A.RgbColorModelHex>()!;
        Assert.Equal("AABBCC", color.Val!.Value);
        C.BarChartSeries[] columns = part.ChartSpace.Descendants<C.BarChartSeries>().ToArray();
        Assert.Equal(2, columns.Length);
        Assert.Empty(columns[1].Elements<C.DataPoint>());
    }

    [Theory]
    [InlineData(OfficeChartKind.Pie)]
    [InlineData(OfficeChartKind.Doughnut)]
    [InlineData(OfficeChartKind.ColumnClustered)]
    public void PointStyles_PreserveNativeStylesAndRenderHatchesAfterReopen(OfficeChartKind kind) {
        var hatch = new OfficeChartPointStyle(OfficeColor.White, hatch: OfficeChartHatchPattern.WideForwardDiagonal,
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
