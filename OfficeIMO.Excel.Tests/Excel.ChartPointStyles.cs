using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
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
            .WithPointStyles(new OfficeChartPointStyle?[] { null, new(noFill: true, outlineColor: OfficeColor.Black), hatch });
        using ExcelDocument authored = ExcelDocument.Create();
        authored.AddWorksheet("Results").AddChart(kind,
            new OfficeChartData(new[] { "Pass", "Unknown", "Fail" }, new[] { series }), 1, 1);
        using var bytes = new MemoryStream(authored.ToBytes());
        using ExcelDocument reopened = ExcelDocument.Load(bytes);
        ExcelChart chart = Assert.Single(reopened.Sheets.Single(sheet => sheet.Name == "Results").Charts);
        Assert.True(chart.TryGetSnapshot(out ExcelChartSnapshot snapshot));
        Assert.True(snapshot.Data.Series[0].PointStyles![1]!.NoFill);
        Assert.Equal(hatch.Hatch, snapshot.Data.Series[0].PointStyles![2]!.Hatch);
        string svg = System.Text.Encoding.UTF8.GetString(chart.ExportImage(OfficeImageExportFormat.Svg).Bytes);
        Assert.Contains("#7300A3", svg, System.StringComparison.OrdinalIgnoreCase);
        Assert.Contains("clipPath", svg, System.StringComparison.Ordinal);
        chart.SetDataPointStyle(0, 1, null).SetDataPointStyle(0, 2, null);
        Assert.True(chart.TryGetSnapshot(out snapshot));
        Assert.Null(snapshot.Data.Series[0].PointStyles);
    }
}
