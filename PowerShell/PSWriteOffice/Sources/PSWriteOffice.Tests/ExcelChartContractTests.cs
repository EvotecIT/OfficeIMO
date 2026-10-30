using System.Data;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel;
using PSWriteOffice.Cmdlets.Excel;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class ExcelChartContractTests {
    [Fact]
    public void SelectedFormattedNumericAndFormulaColumnsPersistInRequestedOrderAtTheRequestedAddress() {
        using var book = ExcelDocument.Create();
        var sheet = book.AddWorksheet("Sales");
        var rows = new DataTable();
        rows.Columns.Add("Revenue", typeof(double));
        rows.Columns.Add("Note", typeof(string));
        rows.Columns.Add("Region", typeof(string));
        rows.Columns.Add("Cost", typeof(double));
        rows.Columns.Add("Forecast", typeof(double));
        rows.Rows.Add(100.125d, "First", "North", 0.125d, 0d);
        rows.Rows.Add(150.875d, "Second", "South", 0.205d, 0d);
        sheet.InsertDataTableAsTable(rows, startRow: 4, startColumn: 3, tableName: "Sales");
        sheet.Range("C5:C6").SetNumberFormat("\"$\"#,##0.00");
        sheet.Range("F5:F6").SetNumberFormat("0.0%");
        sheet.CellFormula(5, 7, "C5*2");
        sheet.CellFormula(6, 7, "C6*2");
        sheet.Range("G5:G6").SetNumberFormat("0.0\" forecast\"");
        book.Calculate();
        sheet.CellValue(20, 3, 999999d);
        using var runspace = Open(book);
        var result = Invoke(runspace, """
            Add-OfficeExcelChart -Document $book -Sheet Sales -Table Sales -Category region -Value Cost,Forecast,Revenue -Address F2 -Width 18cm -Height 9cm -Title 'Margin' -Type BarClustered -Preset CompactComparison -PassThru
            """);
        Assert.IsType<ExcelChart>(Assert.Single(result).BaseObject);
        using var artifact = Save(book);
        using (var saved = ExcelDocument.Load(artifact)) {
            var reopened = saved["Sales"];
            var chart = Assert.Single(reopened.Charts);
            Assert.Equal("Margin", chart.Title);
            Assert.Equal(ExcelChartType.BarClustered, chart.ChartType);
            AssertData(chart, new[] { "North", "South" }, new[] { "Cost", "Forecast", "Revenue" },
                new[] { 0.125d, 0.205d }, new[] { 200.25d, 301.75d }, new[] { 100.125d, 150.875d });
            Assert.Equal("C4:G6", reopened.GetTableRange("Sales"));
            Assert.Equal("\"$\"#,##0.00", reopened.CellAt(5, 3).GetStyle().NumberFormatCode);
            Assert.Equal("0.0%", reopened.CellAt(5, 6).GetStyle().NumberFormatCode);
            Assert.Equal("0.0\" forecast\"", reopened.CellAt(5, 7).GetStyle().NumberFormatCode);
            Assert.Equal("C5*2", reopened.GetFormulaText(5, 7));
            Assert.True(reopened.TryGetCellText(20, 3, out string unrelated));
            Assert.Equal("999999", unrelated);
        }
        AssertPlacement(artifact, row: 1, column: 5, width: 680, height: 340);
        artifact.Position = 0;
        using var package = SpreadsheetDocument.Open(artifact, false);
        var chartXml = Assert.Single(package.WorkbookPart!.WorksheetParts
            .SelectMany(part => part.DrawingsPart?.ChartParts ?? Enumerable.Empty<ChartPart>())).ChartSpace;
        Assert.Empty(chartXml.Descendants<C.Legend>());
        Assert.Contains(chartXml.Descendants<C.ShowValue>(), flag => flag.Val?.Value == true);
    }

    [Fact]
    public void InferredValuesUseOnlyNumericColumnsAndExcludeTotalsAndUnrelatedCells() {
        using var book = ExcelDocument.Create();
        var sheet = book.AddWorksheet("Sales");
        var rows = new DataTable();
        rows.Columns.Add("Region", typeof(string));
        rows.Columns.Add("Revenue", typeof(double));
        rows.Columns.Add("Note", typeof(string));
        rows.Columns.Add("Cost", typeof(double));
        rows.Columns.Add("Empty", typeof(string));
        rows.Columns.Add("Mixed", typeof(string));
        rows.Rows.Add("North", 100d, "First", 40d, DBNull.Value, "5");
        rows.Rows.Add("South", 150d, "Second", 60d, DBNull.Value, "unknown");
        sheet.InsertDataTableAsTable(rows, startRow: 3, startColumn: 2, tableName: "Sales");
        sheet.SetTableTotalsByName("Sales", new Dictionary<string, ExcelTableTotalsFunction> { ["Revenue"] = ExcelTableTotalsFunction.Sum });
        sheet.CellValue(30, 3, 999999d);
        using var runspace = Open(book);
        Assert.Empty(Invoke(runspace, "Add-OfficeExcelChart -Document $book -Table Sales -Address H3"));
        using var artifact = Save(book);
        using var saved = ExcelDocument.Load(artifact);
        AssertData(Assert.Single(saved["Sales"].Charts), new[] { "North", "South" }, new[] { "Revenue", "Cost" }, new[] { 100d, 150d }, new[] { 40d, 60d });
    }

    [Fact]
    public void PipelineRowsWriteOneTableWithoutOverwritingExistingCellsAndReturnOneNativeChart() {
        using var book = ExcelDocument.Create();
        var sheet = book.AddWorksheet("Sales");
        sheet.CellValue(8, 1, "Keep this row");
        using var runspace = Open(book);
        var result = Invoke(runspace, """
            @(
                [pscustomobject]@{ Note = 'First'; Revenue = 100; Region = 'North'; Cost = 40 }
                [pscustomobject]@{ Note = 'Second'; Revenue = 150; Region = 'South'; Cost = 60 }
            ) | Add-OfficeExcelChart -Document $book -Sheet 1 -Category Region -Value Cost,Revenue -Address H2 -Type Line -PassThru
            """);
        Assert.IsType<ExcelChart>(Assert.Single(result).BaseObject);
        using var artifact = Save(book);
        using var saved = ExcelDocument.Load(artifact);
        var reopened = saved["Sales"];
        AssertData(Assert.Single(reopened.Charts), new[] { "North", "South" }, new[] { "Cost", "Revenue" }, new[] { 40d, 60d }, new[] { 100d, 150d });
        var table = Assert.Single(reopened.GetTables());
        Assert.Equal("A10:D12", table.Range);
        Assert.Equal(new[] { "Region", "Cost", "Revenue", "Note" }, table.Columns.Select(column => column.Name));
        Assert.True(reopened.TryGetCellText(8, 1, out string original));
        Assert.Equal("Keep this row", original);
    }

    [Fact]
    public void ExplicitInputTableIsCopiedAndCacheSuppressionKeepsNativeWorksheetReferences() {
        using var book = ExcelDocument.Create();
        book.AddWorksheet("Sales");
        var rows = new DataTable();
        rows.Columns.Add("Revenue", typeof(double));
        rows.Columns.Add("Region", typeof(string));
        rows.Rows.Add(100d, "North");
        using var runspace = Open(book);
        runspace.SessionStateProxy.SetVariable("rows", rows);
        Assert.Empty(Invoke(runspace, "$book | Add-OfficeExcelChart -InputObject $rows -Category Region -Value Revenue -Address F2 -NoCachedData"));
        Assert.Equal(new[] { "Revenue", "Region" }, rows.Columns.Cast<DataColumn>().Select(column => column.ColumnName));
        using var artifact = Save(book);
        using (var saved = ExcelDocument.Load(artifact)) {
            AssertData(Assert.Single(saved["Sales"].Charts), new[] { "North" }, new[] { "Revenue" }, new[] { 100d });
        }
        artifact.Position = 0;
        using var package = SpreadsheetDocument.Open(artifact, false);
        var chart = Assert.Single(package.WorkbookPart!.WorksheetParts.SelectMany(part => part.DrawingsPart?.ChartParts ?? Enumerable.Empty<ChartPart>())).ChartSpace;
        Assert.Empty(chart.Descendants<C.NumberingCache>());
        Assert.Empty(chart.Descendants<C.CategoryAxisData>().SelectMany(category => category.Descendants<C.StringCache>()));
        Assert.Contains(chart.Descendants<C.Formula>(), formula => formula.Text.Contains("$B$2"));
    }

    [Fact]
    public void HeaderlessRangeUsesTheWorkbookPipelineAndReturnsQuietly() {
        using var book = ExcelDocument.Create();
        var sheet = book.AddWorksheet("Sales");
        sheet.CellValue(3, 2, "North");
        sheet.CellValue(3, 3, 100d);
        sheet.CellValue(4, 2, "South");
        sheet.CellValue(4, 3, 150d);
        using var runspace = Open(book);
        Assert.Empty(Invoke(runspace, "$book | Add-OfficeExcelChart -Range B3:C4 -NoHeader -Address E3 -Width 480px -Height 240px"));
        using var artifact = Save(book);
        using (var saved = ExcelDocument.Load(artifact)) {
            var chart = Assert.Single(saved["Sales"].Charts);
            Assert.True(chart.TryGetData(out var data));
            Assert.Equal(new[] { "North", "South" }, data.Categories);
            Assert.Equal(new[] { 100d, 150d }, Assert.Single(data.Series).Values);
        }
        AssertPlacement(artifact, row: 2, column: 4, width: 480, height: 240);
    }

    [Fact]
    public void TableChartResolvesTheActiveWorksheetCompositionScope() {
        using var book = ExcelDocument.Create();
        var sheet = book.AddWorksheet("Sales");
        SeedSales(sheet);
        using var runspace = Open(book);
        Assert.Empty(Invoke(runspace, """
            Add-OfficeExcelSheet -Document $book -Name Sales {
                Add-OfficeExcelChart -Table Sales -Category Region -Value Revenue -Address F2
            }
            """));
        using var artifact = Save(book);
        using var saved = ExcelDocument.Load(artifact);
        AssertData(Assert.Single(saved["Sales"].Charts), new[] { "North", "South" }, new[] { "Revenue" }, new[] { 100d, 150d });
    }

    [Theory]
    [InlineData("-Table Missing")]
    [InlineData("-Table Sales -Category Missing")]
    [InlineData("-Table Sales -Value Missing")]
    [InlineData("-Table Sales -Value Note")]
    [InlineData("-Table Sales -Value Revenue,revenue")]
    [InlineData("-Table Sales -Category Revenue -Value Revenue")]
    [InlineData("-Table Sales -Value Revenue,Cost -Type Pie")]
    [InlineData("-Table Sales -Type Scatter")]
    [InlineData("-Range A:B")]
    [InlineData("-Range A1:D3")]
    public void InvalidSelectionsLeaveTheWorkbookUnchanged(string selection) {
        using var book = ExcelDocument.Create();
        var sheet = book.AddWorksheet("Sales");
        SeedSales(sheet);
        string originalRange = sheet.UsedRangeA1;
        int sheetCount = book.Sheets.Count;
        using var runspace = Open(book);
        Reject(runspace, "Add-OfficeExcelChart -Document $book -Address F2 " + selection);
        Assert.Empty(sheet.Charts);
        Assert.Equal(sheetCount, book.Sheets.Count);
        Assert.Equal(originalRange, sheet.UsedRangeA1);
        Assert.Equal("A1:D3", Assert.Single(sheet.GetTables()).Range);
        Assert.Equal(100d, sheet.CellAt(2, 2).GetValue<double>());
    }

    [Fact]
    public void InferenceDoesNotInventAValueSeriesForATextOnlyTable() {
        using var book = ExcelDocument.Create();
        var sheet = book.AddWorksheet("Sales");
        var rows = new DataTable();
        rows.Columns.Add("Region", typeof(string));
        rows.Columns.Add("Note", typeof(string));
        rows.Rows.Add("North", "No numeric values");
        sheet.InsertDataTableAsTable(rows, tableName: "Sales");
        using var runspace = Open(book);
        Reject(runspace, "Add-OfficeExcelChart -Document $book -Table Sales -Address F2");
        Assert.Empty(sheet.Charts);
        Assert.Equal("A1:B2", Assert.Single(sheet.GetTables()).Range);
    }

    [Theory]
    [InlineData("-Category Missing -Value Revenue -Address F2")]
    [InlineData("-Category Region -Value Missing -Address F2")]
    [InlineData("-Category Region -Value Revenue,revenue -Address F2")]
    [InlineData("-Category Region -Value Revenue -Address F2:F3")]
    [InlineData("-Category Region -Value Revenue -Address F2 -Width 0px")]
    [InlineData("-Category Region -Value Revenue -Address F2 -Height -2cm")]
    [InlineData("-Category Region -Value Revenue -Address F2 -Type Bubble")]
    public void InvalidInputMappingOrPlacementDoesNotWriteATable(string parameters) {
        using var book = ExcelDocument.Create();
        var sheet = book.AddWorksheet("Sales");
        using var runspace = Open(book);
        Reject(runspace, "Add-OfficeExcelChart -Document $book -InputObject @([pscustomobject]@{ Region = 'North'; Revenue = 100 }) " + parameters);
        Assert.Empty(sheet.GetTables());
        Assert.Empty(sheet.Charts);
        Assert.False(sheet.TryGetCellText(1, 1, out _));
    }

    [Theory]
    [InlineData("bad")]
    [InlineData("NaN")]
    [InlineData("Infinity")]
    public void InvalidLaterPipelineValueDoesNotWriteEarlierRows(string value) {
        using var book = ExcelDocument.Create();
        var sheet = book.AddWorksheet("Sales");
        using var runspace = Open(book);
        runspace.SessionStateProxy.SetVariable("invalid", value);
        Reject(runspace, """
            @(
                [pscustomobject]@{ Region = 'North'; Revenue = 100 }
                [pscustomobject]@{ Region = 'South'; Revenue = $invalid }
            ) | Add-OfficeExcelChart -Document $book -Category Region -Value Revenue -Address F2
            """);
        Assert.Empty(sheet.GetTables());
        Assert.Empty(sheet.Charts);
        Assert.False(sheet.TryGetCellText(1, 1, out _));
    }

    private static void SeedSales(ExcelSheet sheet) {
        var rows = new DataTable();
        rows.Columns.Add("Region", typeof(string));
        rows.Columns.Add("Revenue", typeof(double));
        rows.Columns.Add("Note", typeof(string));
        rows.Columns.Add("Cost", typeof(double));
        rows.Rows.Add("North", 100d, "First", 40d);
        rows.Rows.Add("South", 150d, "Second", 60d);
        sheet.InsertDataTableAsTable(rows, tableName: "Sales");
    }

    private static void AssertData(ExcelChart chart, string[] categories, string[] names, params double[][] values) {
        Assert.True(chart.TryGetData(out var data));
        Assert.Equal(categories, data.Categories);
        Assert.Equal(names, data.Series.Select(series => series.Name));
        Assert.Equal(values.Length, data.Series.Count);
        for (int i = 0; i < values.Length; i++) Assert.Equal(values[i], data.Series[i].Values);
    }

    private static void AssertPlacement(MemoryStream artifact, int row, int column, int width, int height) {
        artifact.Position = 0;
        using var package = SpreadsheetDocument.Open(artifact, false);
        var anchor = Assert.Single(package.WorkbookPart!.WorksheetParts
            .SelectMany(part => part.DrawingsPart?.WorksheetDrawing?.Elements<Xdr.OneCellAnchor>() ?? Enumerable.Empty<Xdr.OneCellAnchor>()));
        Assert.Equal(row.ToString(System.Globalization.CultureInfo.InvariantCulture), anchor.FromMarker!.RowId!.Text);
        Assert.Equal(column.ToString(System.Globalization.CultureInfo.InvariantCulture), anchor.FromMarker!.ColumnId!.Text);
        Assert.Equal(width * 9525L, anchor.Extent!.Cx!.Value);
        Assert.Equal(height * 9525L, anchor.Extent!.Cy!.Value);
    }

    private static MemoryStream Save(ExcelDocument book) {
        var artifact = new MemoryStream();
        book.Save(artifact);
        artifact.Position = 0;
        return artifact;
    }

    private static Runspace Open(ExcelDocument book) {
        var state = InitialSessionState.CreateDefault();
        state.Commands.Add(new SessionStateCmdletEntry("Add-OfficeExcelChart", typeof(AddOfficeExcelChartCommand), null));
        state.Commands.Add(new SessionStateCmdletEntry("Add-OfficeExcelSheet", typeof(AddOfficeExcelSheetCommand), null));
        var runspace = RunspaceFactory.CreateRunspace(state);
        runspace.Open();
        runspace.SessionStateProxy.SetVariable("book", book);
        return runspace;
    }

    private static System.Collections.ObjectModel.Collection<PSObject> Invoke(Runspace runspace, string script) {
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript(script);
        var output = shell.Invoke();
        Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
        return output;
    }

    private static void Reject(Runspace runspace, string script) {
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript(script);
        var failure = Record.Exception(() => shell.Invoke());
        Assert.True(failure != null || shell.HadErrors);
    }
}
