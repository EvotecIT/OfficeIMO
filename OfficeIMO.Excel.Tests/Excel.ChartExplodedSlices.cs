using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class ExcelChartExplodedSlicesTests {
    [Fact]
    public void ExcelProducedPieAndDoughnut_ProjectExplodedPoint() {
        string path = Path.Combine(AppContext.BaseDirectory, "Documents", "Charts", "Excel",
            "exploded-slices.xlsx");
        using ExcelDocument document = ExcelDocument.Load(path);
        ExcelChart[] charts = document.Sheets.Single(sheet => sheet.Name == "Exploded")
            .Charts.ToArray();
        Assert.Equal(2, charts.Length);
        foreach (ExcelChart chart in charts) {
            Assert.True(chart.TryGetSnapshot(out ExcelChartSnapshot snapshot));
            Assert.Equal(new[] { 25, 0 }, snapshot.Data.Series.Single().PointExplosions);
            Assert.NotEmpty(chart.ExportImage(OfficeImageExportFormat.Svg).Bytes);
        }
    }

    [Theory]
    [InlineData(OfficeChartKind.Pie)]
    [InlineData(OfficeChartKind.Doughnut)]
    public void ExplodedPoint_RoundTripsIntoStaticExportAndSurvivesValueUpdate(OfficeChartKind kind) {
        var authored = new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeChartSeries("Status", new[] { 7d, 3d }).WithPointExplosions(new[] { 25, 0 })
        });
        using var document = ExcelDocument.Create();
        document.AddWorksheet("Results").AddChart(kind, authored, 1, 1);
        using var bytes = new MemoryStream(document.ToBytes());
        using ExcelDocument reopened = ExcelDocument.Load(bytes);
        ExcelChart chart = reopened.Sheets.Single(sheet => sheet.Name == "Results").Charts.Single();
        Assert.True(chart.TryGetSnapshot(out ExcelChartSnapshot snapshot));
        Assert.Equal(new[] { 25, 0 }, snapshot.Data.Series.Single().PointExplosions);
        Assert.NotEmpty(chart.ExportImage(OfficeImageExportFormat.Svg).Bytes);
        chart.UpdateData(new ExcelChartData(new[] { "A", "B" }, new[] {
            new ExcelChartSeries("Status", new[] { 8d, 2d })
        }));
        C.PieChartSeries native = reopened.OpenXmlDocument.WorkbookPart!.WorksheetParts
            .Single(part => part.DrawingsPart != null).DrawingsPart!.ChartParts.Single()
            .ChartSpace!.Descendants<C.PieChartSeries>().Single();
        Assert.Equal((uint)25, native.Elements<C.DataPoint>()
            .Single(point => point.Index!.Val!.Value == 0).GetFirstChild<C.Explosion>()!.Val!.Value);
        Assert.Empty(reopened.ValidateDocument());
    }

    [Fact]
    public void UnsupportedNativeExplosion_RejectsStaticSnapshotWithoutChangingNativeXml() {
        using var document = ExcelDocument.Create();
        ExcelChart chart = document.AddWorksheet("Results").AddChart(OfficeChartKind.Pie,
            new OfficeChartData(new[] { "A", "B" }, new[] {
                new OfficeChartSeries("Status", new[] { 7d, 3d })
            }), 1, 1);
        C.PieChartSeries native = document.OpenXmlDocument.WorkbookPart!.WorksheetParts
            .Single(part => part.DrawingsPart != null).DrawingsPart!.ChartParts.Single()
            .ChartSpace!.Descendants<C.PieChartSeries>().Single();
        native.AddChild(new C.Explosion { Val = 401U }, true);
        Assert.False(chart.TryGetSnapshot(out _));
        Assert.Equal((uint)401, native.GetFirstChild<C.Explosion>()!.Val!.Value);
    }
}
