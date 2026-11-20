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

    [Fact]
    public void AutomaticTitleUsesTheCurrentSeriesNameCell() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = Sheet(document);
        ExcelChart chart = sheet.AddChart(OfficeChartType.ColumnClustered, "A1:B4",
            new OfficeIMO.Excel.ExcelChartOptions { Row = 1, Column = 4 });
        C.Chart native = Assert.Single(sheet.WorksheetPart.DrawingsPart!.ChartParts).ChartSpace!.GetFirstChild<C.Chart>()!;
        native.PrependChild(new C.Title(new C.Overlay { Val = false }));
        native.GetFirstChild<C.AutoTitleDeleted>()!.Val = false;
        sheet.Cell(1, 2).SetValue("Forecast");

        Assert.True(chart.TryGetSnapshot(out ExcelChartInfo snapshot));

        Assert.Equal("Forecast", snapshot.Title);
    }

    [Fact]
    public void SnapshotKeepsEmptyCellsMissingInRowsWaitingForTheWriter() {
        var table = new System.Data.DataTable();
        table.Columns.Add("Month", typeof(object));
        table.Columns.Add("Actual", typeof(object));
        table.Rows.Add("Jan", 120);
        table.Rows.Add("Feb", null);
        table.Rows.Add("Mar", 160);
        using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = document.AddSheet("Data");
        sheet.AddTable(table, new ExcelTableOptions { Name = "Scores", AutoFit = false }, row: 1, column: 1);
        ExcelChart chart = sheet.AddChart(OfficeChartType.Line, "A1:B4", new OfficeIMO.Excel.ExcelChartOptions { Row = 1, Column = 4 });

        // The rows still wait for the writer, so the chart data comes from the deferred reader.
        Assert.True(document.HasUnmaterializedDirectDataSetRows);
        Assert.True(chart.TryGetSnapshot(out ExcelChartInfo snapshot));

        Assert.Equal(new[] { 120D, double.NaN, 160D }, snapshot.Data.Series.Single().Values);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void SnapshotLeavesOutHiddenRowsWhenThePlotShowsVisibleCellsOnly(bool visibleOnly) {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = Sheet(document);
        sheet.Cell(3, 2).SetValue(140);
        sheet.SetRowHidden(3, true);
        ExcelChart chart = sheet.AddChart(OfficeChartType.Line, "A1:B4", new OfficeIMO.Excel.ExcelChartOptions { Row = 1, Column = 4 });
        ChartPart part = Assert.Single(sheet.WorksheetPart.DrawingsPart!.ChartParts);
        part.ChartSpace!.GetFirstChild<C.Chart>()!.PlotVisibleOnly!.Val = visibleOnly;

        Assert.True(chart.TryGetSnapshot(out ExcelChartInfo snapshot));

        Assert.Equal(visibleOnly ? new[] { "Jan", "Mar" } : new[] { "Jan", "Feb", "Mar" }, snapshot.Data.Categories);
        Assert.Equal(visibleOnly ? new[] { 120D, 160D } : new[] { 120D, 140D, 160D }, snapshot.Data.Series.Single().Values);
    }

    [Fact]
    public void SnapshotDrawsLogarithmicAndReversedValueAxes() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = Sheet(document);
        ExcelChart chart = sheet.AddChart(OfficeChartType.ColumnClustered, "A1:B4", new OfficeIMO.Excel.ExcelChartOptions { Row = 1, Column = 4 });
        ChartPart part = Assert.Single(sheet.WorksheetPart.DrawingsPart!.ChartParts);
        C.Scaling scaling = part.ChartSpace!.Descendants<C.ValueAxis>().Single().Scaling!;
        scaling.PrependChild(new C.LogBase { Val = 10 });
        scaling.GetFirstChild<C.Orientation>()!.Val = C.OrientationValues.MaxMin;

        Assert.True(chart.TryGetSnapshot(out ExcelChartInfo snapshot));

        Assert.Equal(10D, snapshot.Layout!.VerticalAxisLogBase);
        Assert.True(snapshot.Layout.ReverseVerticalValueAxis);
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
