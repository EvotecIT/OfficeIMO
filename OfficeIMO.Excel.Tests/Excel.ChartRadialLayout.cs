using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelChartRadialLayoutTests {
    [Theory]
    [InlineData(OfficeChartKind.Pie)]
    [InlineData(OfficeChartKind.Doughnut)]
    public void RadialGeometry_PersistsThroughReopenUpdateAndSnapshot(OfficeChartKind kind) {
        var data = new OfficeChartData(new[] { "A", "B" }, new[] { new OfficeChartSeries("Status", new[] { 7d, 3d }) });
        using ExcelDocument authored = ExcelDocument.Create();
        authored.AddWorksheet("Results").AddChart(kind, data, 1, 1).SetRadialLayout(new OfficeChartRadialLayout(135, 75));
        using var bytes = new MemoryStream(authored.ToBytes());
        using ExcelDocument reopened = ExcelDocument.Load(bytes);
        ExcelChart chart = reopened.Sheets.Single(sheet => sheet.Name == "Results").Charts.Single();
        chart.UpdateData(new ExcelChartData(new[] { "A", "B" }, new[] { new ExcelChartSeries("Status", new[] { 8d, 2d }) }));
        Assert.True(chart.TryGetSnapshot(out ExcelChartSnapshot snapshot));
        Assert.Equal(135, snapshot.RadialLayout.FirstSliceAngleDegrees);
        Assert.Equal(kind == OfficeChartKind.Doughnut ? 75 : 50, snapshot.RadialLayout.DoughnutHolePercent);
        Assert.Empty(reopened.ValidateDocument());
    }
}
