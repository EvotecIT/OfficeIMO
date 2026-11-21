using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Tests;

/// <summary>Pins how worksheet charts hand empty cells, hidden cells and automatic titles to the shared renderer.</summary>
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
    public void SnapshotLeavesOutHiddenCellsOfEachSeriesRangeAsExcelDoes() {
        // Excel 16 closes the gap in each range on its own: the first series loses Feb, the second
        // keeps its four values, and the fourth category has no label.
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddSheet("Data");
        string[] months = { "Jan", "Feb", "Mar", "Apr" };
        for (int index = 0; index < months.Length; index++) {
            sheet.Cell(2 + index, 1).SetValue(months[index]);
            sheet.Cell(2 + index, 2).SetValue(10 * (index + 1));
            sheet.Cell(7 + index, 4).SetValue(100 * (index + 1));
        }
        sheet.SetRowHidden(3, true);
        ExcelChart chart = sheet.AddChartFromRanges(OfficeChartType.ColumnClustered, new[] {
            new ExcelChartSeriesRange("S1", "A2:A5", "B2:B5"), new ExcelChartSeriesRange("S2", "A2:A5", "D7:D10") },
            new OfficeIMO.Excel.ExcelChartOptions { Row = 1, Column = 6 });

        Assert.True(chart.TryGetSnapshot(out ExcelChartInfo snapshot));

        Assert.Equal(new[] { "Jan", "Mar", "Apr", "" }, snapshot.Data.Categories);
        Assert.Equal(new[] { 10D, 30D, 40D, double.NaN }, snapshot.Data.Series[0].Values);
        Assert.Equal(new[] { 100D, 200D, 300D, 400D }, snapshot.Data.Series[1].Values);
    }

    [Fact]
    public void SnapshotLeavesOutASeriesWhoseColumnIsHiddenAndKeepsTheOthersFormatting() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = Sheet(document);
        sheet.Cell(1, 3).SetValue("Plan");
        sheet.Cell(2, 3).SetValue(100);
        sheet.Cell(3, 3).SetValue(110);
        sheet.Cell(4, 3).SetValue(130);
        ExcelChart chart = sheet.AddChart(OfficeChartType.ColumnClustered, "A1:C4", new OfficeIMO.Excel.ExcelChartOptions { Row = 1, Column = 5 });
        C.BarChartSeries plan = Assert.Single(sheet.WorksheetPart.DrawingsPart!.ChartParts).ChartSpace!.Descendants<C.BarChartSeries>().ElementAt(1);
        plan.RemoveAllChildren<C.ChartShapeProperties>();
        plan.InsertAfter(new C.ChartShapeProperties(new A.SolidFill(new A.RgbColorModelHex { Val = "FF0000" })), plan.GetFirstChild<C.SeriesText>());
        sheet.SetColumnHidden(2, true);

        Assert.True(chart.TryGetSnapshot(out ExcelChartInfo snapshot));

        ExcelChartSeries series = Assert.Single(snapshot.Data.Series);
        Assert.Equal("Plan", series.Name);
        Assert.Equal(new[] { 100D, 110D, 130D }, series.Values);
        Assert.Equal("FF0000", series.SeriesColorArgb?.Substring(series.SeriesColorArgb.Length - 6));
    }

    [Fact]
    public void SnapshotLeavesCategoryLabelsEmptyWhenOnlyTheirColumnIsHidden() {
        // Excel 16 draws the three visible values with empty category labels.
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddSheet("Data");
        for (int index = 0; index < 3; index++) {
            sheet.Cell(2 + index, 1).SetValue("Cat" + (index + 2));
            sheet.Cell(7 + index, 4).SetValue(10 * (index + 7));
        }
        sheet.SetColumnHidden(1, true);
        ExcelChart chart = sheet.AddChartFromRanges(OfficeChartType.ColumnClustered, new[] {
            new ExcelChartSeriesRange("S1", "A2:A4", "D7:D9") }, new OfficeIMO.Excel.ExcelChartOptions { Row = 1, Column = 6 });

        Assert.True(chart.TryGetSnapshot(out ExcelChartInfo snapshot));

        Assert.Equal(new[] { "", "", "" }, snapshot.Data.Categories);
        Assert.Equal(new[] { 70D, 80D, 90D }, snapshot.Data.Series.Single().Values);
    }

    [Fact]
    public void SnapshotTakesCategoryLabelsFromTheFirstSeriesStillPlotted() {
        // With the first series hidden, Excel 16 labels the axis from the next series' own category cells.
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddSheet("Data");
        string[] months = { "Jan", "Feb", "Mar" }, letters = { "X", "Y", "Z" };
        for (int index = 0; index < 3; index++) {
            sheet.Cell(2 + index, 1).SetValue(months[index]);
            sheet.Cell(2 + index, 2).SetValue(10 * (index + 1));
            sheet.Cell(6 + index, 1).SetValue(letters[index]);
            sheet.Cell(6 + index, 3).SetValue(100 * (index + 1));
        }
        sheet.SetColumnHidden(2, true);
        ExcelChart chart = sheet.AddChartFromRanges(OfficeChartType.ColumnClustered, new[] {
            new ExcelChartSeriesRange("S1", "A2:A4", "B2:B4"), new ExcelChartSeriesRange("S2", "A2:A4", "C6:C8") },
            new OfficeIMO.Excel.ExcelChartOptions { Row = 1, Column = 6 });
        // Office files can give each series its own category cells; point the second series at X, Y and Z.
        C.CategoryAxisData second = Assert.Single(sheet.WorksheetPart.DrawingsPart!.ChartParts).ChartSpace!
            .Descendants<C.BarChartSeries>().ElementAt(1).GetFirstChild<C.CategoryAxisData>()!;
        second.Descendants<C.Formula>().Single().Text = "Data!$A$6:$A$8";
        C.StringPoint[] points = second.Descendants<C.StringPoint>().OrderBy(point => point.Index!.Value).ToArray();
        for (int index = 0; index < points.Length; index++) points[index].NumericValue!.Text = letters[index];

        Assert.True(chart.TryGetSnapshot(out ExcelChartInfo snapshot));

        Assert.Equal(letters, snapshot.Data.Categories);
        Assert.Equal(new[] { 100D, 200D, 300D }, Assert.Single(snapshot.Data.Series).Values);
    }

    [Fact]
    public void ChartWhoseEverySeriesIsHiddenHasNoSnapshot() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = Sheet(document);
        ExcelChart chart = sheet.AddChart(OfficeChartType.ColumnClustered, "A1:B4", new OfficeIMO.Excel.ExcelChartOptions { Row = 1, Column = 4 });
        sheet.SetColumnHidden(2, true);

        // Excel plots nothing, so the hidden values must not come back as a chart.
        Assert.False(chart.TryGetSnapshot(out _));
    }

    [Fact]
    public void SnapshotLeavesOutRowsWithoutRecordsWhenTheSheetHidesRowsByDefault() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = Sheet(document);
        sheet.Cell(5, 1).SetValue("Apr");
        sheet.Cell(5, 2).SetValue(170);
        ExcelChart chart = sheet.AddChart(OfficeChartType.Line, "A1:B5", new OfficeIMO.Excel.ExcelChartOptions { Row = 1, Column = 4 });
        S.Worksheet worksheet = sheet.WorksheetPart.Worksheet!;
        // Under a zero-height default a row without a record is hidden; rows with records stay visible.
        worksheet.Descendants<S.Row>().Single(row => row.RowIndex?.Value == 4U).Remove();
        S.SheetFormatProperties format = worksheet.GetFirstChild<S.SheetFormatProperties>() ??
            worksheet.InsertBefore(new S.SheetFormatProperties { DefaultRowHeight = 15D },
                (DocumentFormat.OpenXml.OpenXmlElement?)worksheet.GetFirstChild<S.Columns>() ?? worksheet.GetFirstChild<S.SheetData>());
        format.ZeroHeight = true;

        Assert.True(chart.TryGetSnapshot(out ExcelChartInfo snapshot));

        Assert.Equal(new[] { "Jan", "Feb", "Apr" }, snapshot.Data.Categories);
        Assert.Equal(new[] { 120D, double.NaN, 170D }, snapshot.Data.Series.Single().Values);
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

    [Fact]
    public void CustomErrorAmountsFollowTheVisiblePointsTheyBelongTo() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = Sheet(document);
        sheet.Cell(3, 2).SetValue(140);
        sheet.Cell(2, 3).SetValue(10);
        sheet.Cell(3, 3).SetValue(20);
        sheet.Cell(4, 3).SetValue(90);
        sheet.SetRowHidden(3, true);
        ExcelChart chart = sheet.AddChart(OfficeChartType.Line, "A1:B4", new OfficeIMO.Excel.ExcelChartOptions { Row = 1, Column = 5 });
        C.LineChartSeries series = Assert.Single(sheet.WorksheetPart.DrawingsPart!.ChartParts).ChartSpace!.Descendants<C.LineChartSeries>().Single();
        series.InsertBefore(new C.ErrorBars(new C.ErrorDirection { Val = C.ErrorBarDirectionValues.Y },
                new C.ErrorBarType { Val = C.ErrorBarValues.Plus }, new C.ErrorBarValueType { Val = C.ErrorValues.Custom },
                new C.NoEndCap { Val = false },
                new C.Plus(new C.NumberReference(new C.Formula("Data!$C$2:$C$4"), new C.NumberingCache(new C.FormatCode("General"),
                    new C.PointCount { Val = 3U },
                    new C.NumericPoint(new C.NumericValue("10")) { Index = 0U },
                    new C.NumericPoint(new C.NumericValue("20")) { Index = 1U },
                    new C.NumericPoint(new C.NumericValue("90")) { Index = 2U })))),
            series.GetFirstChild<C.CategoryAxisData>());

        string svg = System.Text.Encoding.UTF8.GetString(chart.ToImage(OfficeImageExportFormat.Svg).Bytes);

        // March (160) keeps its own +90 once February's row is hidden, so the value axis reaches 250.
        double largest = System.Text.RegularExpressions.Regex.Matches(svg, ">([0-9]+)<").Cast<System.Text.RegularExpressions.Match>()
            .Select(match => double.Parse(match.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture)).Max();
        Assert.True(largest >= 250D, $"The value axis reaches {largest}.");
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
