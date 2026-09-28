using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class ExcelChartRadialOutsideLabelsTests {
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
}
