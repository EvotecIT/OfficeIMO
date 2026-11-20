using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

/// <summary>Pins how worksheet charts hand empty cells and automatic titles to the shared renderer.</summary>
public sealed class ExcelChartPlotSemanticsTests {
    [Fact]
    public void SnapshotKeepsEmptyValueCellsMissing() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = Sheet(document);
        ExcelChart chart = sheet.AddChart(OfficeChartType.Line, "A1:B4", new OfficeIMO.Excel.ExcelChartOptions { Row = 1, Column = 4 });

        Assert.True(chart.TryGetSnapshot(out ExcelChartInfo snapshot));

        Assert.Equal(new[] { 120D, double.NaN, 160D }, snapshot.Data.Series.Single().Values);
    }

    [Fact]
    public void SnapshotTitlesASingleSeriesChartWithItsSeriesName() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = Sheet(document);
        ExcelChart chart = sheet.AddChart(OfficeChartType.ColumnClustered, "A1:B4",
            new OfficeIMO.Excel.ExcelChartOptions { Row = 1, Column = 4 });
        ChartPart part = Assert.Single(sheet.WorksheetPart.DrawingsPart!.ChartParts);
        C.Chart native = part.ChartSpace!.GetFirstChild<C.Chart>()!;
        native.GetFirstChild<C.Title>()?.Remove();
        native.PrependChild(new C.Title(new C.Overlay { Val = false }));
        native.GetFirstChild<C.AutoTitleDeleted>()!.Val = false;

        Assert.True(chart.TryGetSnapshot(out ExcelChartInfo snapshot));

        Assert.Equal("Actual", snapshot.Title);
    }

    private static ExcelSheet Sheet(ExcelDocument document) {
        ExcelSheet sheet = document.AddSheet("Data");
        sheet.Cell(1, 1).SetValue("Month");
        sheet.Cell(1, 2).SetValue("Actual");
        sheet.Cell(2, 1).SetValue("Jan");
        sheet.Cell(2, 2).SetValue(120);
        sheet.Cell(3, 1).SetValue("Feb");
        sheet.Cell(4, 1).SetValue("Mar");
        sheet.Cell(4, 2).SetValue(160);
        return sheet;
    }
}
