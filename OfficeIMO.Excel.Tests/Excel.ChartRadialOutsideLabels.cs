using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class ExcelChartRadialOutsideLabelsTests {
    [Fact]
    public void OuterRingOnlyLabelsKeepTheirSeriesSelectionAndLeaders() {
        using var document = ExcelDocument.Create();
        ExcelChart chart = document.AddWorksheet("Results").AddChart(OfficeChartKind.Doughnut,
            new OfficeChartData(new[] { "A", "B" }, new[] {
                new OfficeChartSeries("Inner", new[] { 3d, 2d }),
                new OfficeChartSeries("Outer", new[] { 4d, 1d })
            }), 1, 1);
        C.DoughnutChart native = document.OpenXmlDocument.WorkbookPart!.WorksheetParts
            .Single(part => part.DrawingsPart != null).DrawingsPart!.ChartParts.Single()
            .ChartSpace!.Descendants<C.DoughnutChart>().Single();
        native.Elements<C.PieChartSeries>().Last().AddChild(new C.DataLabels(
            new C.DataLabelPosition { Val = C.DataLabelPositionValues.OutsideEnd },
            new C.ShowCategoryName { Val = true },
            new C.ShowLeaderLines { Val = true }), true);

        Assert.True(chart.TryGetSnapshot(out ExcelChartSnapshot snapshot));
        Assert.Equal(new[] { 1 }, snapshot.Layout!.DataLabelSeriesIndexes);
        Assert.True(snapshot.Layout.ShowDataLabelLeaderLines);
        Assert.NotEmpty(chart.ExportImage(OfficeImageExportFormat.Svg).Bytes);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NativeOutsideLabelsRetainLeaderLineSettingInImageSnapshot(bool showLeaderLines) {
        using var document = ExcelDocument.Create();
        ExcelChart chart = document.AddWorksheet("Results").AddChart(OfficeChartKind.Doughnut,
            new OfficeChartData(new[] { "A", "B", "C" }, new[] {
                new OfficeChartSeries("Status", new[] { 3d, 4d, 5d })
            }), 1, 1);
        C.DoughnutChart native = document.OpenXmlDocument.WorkbookPart!.WorksheetParts
            .Single(part => part.DrawingsPart != null).DrawingsPart!.ChartParts.Single()
            .ChartSpace!.Descendants<C.DoughnutChart>().Single();
        var labels = new C.DataLabels();
        labels.AddChild(new C.DataLabelPosition { Val = C.DataLabelPositionValues.OutsideEnd }, true);
        labels.AddChild(new C.ShowCategoryName { Val = true }, true);
        labels.AddChild(new C.ShowLeaderLines { Val = showLeaderLines }, true);
        native.AddChild(labels, true);

        Assert.True(chart.TryGetSnapshot(out ExcelChartSnapshot snapshot));
        Assert.Equal(OfficeChartDataLabelPosition.OutsideEnd, snapshot.Layout!.DataLabelPosition);
        Assert.Equal(showLeaderLines, snapshot.Layout.ShowDataLabelLeaderLines);
        Assert.NotEmpty(chart.ExportImage(OfficeImageExportFormat.Png).Bytes);
    }

    [Fact]
    public void EmptyLeaderLinesContainerDoesNotReportMissingRenderedLines() {
        using var document = ExcelDocument.Create();
        ExcelChart chart = document.AddWorksheet("Results").AddChart(OfficeChartKind.Pie,
            new OfficeChartData(new[] { "A", "B" }, new[] {
                new OfficeChartSeries("Status", new[] { 3d, 4d })
            }), 1, 1);
        C.DataLabels labels = document.OpenXmlDocument.WorkbookPart!.WorksheetParts
            .Single(part => part.DrawingsPart != null).DrawingsPart!.ChartParts.Single()
            .ChartSpace!.Descendants<C.PieChart>().Single().GetFirstChild<C.DataLabels>()!;
        labels.AddChild(new C.DataLabelPosition { Val = C.DataLabelPositionValues.OutsideEnd }, true);
        labels.AddChild(new C.ShowCategoryName { Val = true }, true);
        labels.AddChild(new C.ShowLeaderLines { Val = true }, true);
        labels.AddChild(new C.LeaderLines(), true);

        var result = chart.ExportImage(OfficeImageExportFormat.Png);
        Assert.NotEmpty(result.Bytes);
        Assert.DoesNotContain(result.Diagnostics, item =>
            item.Code == ExcelImageExportDiagnosticCodes.ChartDataLabelLeaderLinesUnsupported);

        labels.GetFirstChild<C.LeaderLines>()!.AddChild(new C.ChartShapeProperties(new A.Outline()), true);
        result = chart.ExportImage(OfficeImageExportFormat.Png);
        Assert.Contains(result.Diagnostics, item =>
            item.Code == ExcelImageExportDiagnosticCodes.ChartDataLabelLeaderLinesUnsupported);
    }

    [Fact]
    public void ConflictingSeriesLeaderLineSettingsRejectImageSnapshot() {
        using var document = ExcelDocument.Create();
        ExcelChart chart = document.AddWorksheet("Results").AddChart(OfficeChartKind.Doughnut,
            new OfficeChartData(new[] { "A", "B" }, new[] {
                new OfficeChartSeries("Inner", new[] { 3d, 4d }),
                new OfficeChartSeries("Outer", new[] { 4d, 3d })
            }), 1, 1);
        C.DoughnutChart native = document.OpenXmlDocument.WorkbookPart!.WorksheetParts
            .Single(part => part.DrawingsPart != null).DrawingsPart!.ChartParts.Single()
            .ChartSpace!.Descendants<C.DoughnutChart>().Single();
        C.PieChartSeries[] series = native.Elements<C.PieChartSeries>().ToArray();
        for (int index = 0; index < series.Length; index++) {
            var labels = new C.DataLabels();
            labels.AddChild(new C.ShowCategoryName { Val = true }, true);
            labels.AddChild(new C.ShowLeaderLines { Val = index == 0 }, true);
            series[index].AddChild(labels, true);
        }
        Assert.False(chart.TryGetSnapshot(out _));
    }
}
