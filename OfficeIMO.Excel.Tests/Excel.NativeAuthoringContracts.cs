using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO;
using OfficeIMO.Excel;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using S = DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Tests;

public sealed class ExcelNativeAuthoringContracts {
    [Theory]
    [InlineData(OfficeChartType.Column3DStacked100, "bar3DChart")]
    [InlineData(OfficeChartType.Line3D, "line3DChart")]
    [InlineData(OfficeChartType.Pie3D, "pie3DChart")]
    public void NativeKindsRemainEditableWithExactPhysicalAnchor(OfficeChartType type, string element) {
        using var workbook = ExcelDocument.Create();
        ExcelSheet sheet = workbook.AddSheet("Summary");
        var options = new ExcelChartOptions {
            Row = 7, Column = 4, Width = OfficeLength.Millimeters(125),
            Height = OfficeLength.Millimeters(75), Title = "Native"
        };
        ExcelChart chart = sheet.AddChart(type, SampleData(), options);
        Assert.Contains(sheet.Charts, candidate => candidate.Name == chart.Name);
        using var output = Save(workbook);
        using var package = SpreadsheetDocument.Open(output, false);
        WorksheetPart part = ChartSheet(package);
        C.ChartSpace space = Assert.Single(part.DrawingsPart!.ChartParts).ChartSpace;
        Assert.Contains(space.Descendants(), node => node.LocalName == element);
        Assert.NotEmpty(space.Descendants<C.Formula>());
        Xdr.OneCellAnchor anchor = Assert.Single(part.DrawingsPart.WorksheetDrawing!.Elements<Xdr.OneCellAnchor>());
        Assert.Equal("6", anchor.FromMarker!.RowId!.Text);
        Assert.Equal("3", anchor.FromMarker.ColumnId!.Text);
        Assert.Equal(options.Width!.Value.ToEmus(), anchor.Extent!.Cx!.Value);
        Assert.Equal(options.Height!.Value.ToEmus(), anchor.Extent.Cy!.Value);
        Assert.Empty(part.DrawingsPart.ImageParts);
        Assert.Empty(new OpenXmlValidator().Validate(package));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void A1SourceRetainsLinksWithAndWithoutCaches(bool cache) {
        using var workbook = ExcelDocument.Create();
        ExcelSheet sheet = workbook.AddSheet("Native Data");
        using DataTable rows = SampleTable();
        sheet.WriteData(rows);
        sheet.AddChart(OfficeChartType.Line, "A1:B3", new ExcelChartOptions { IncludeCachedData = cache });
        using var output = Save(workbook);
        using var package = SpreadsheetDocument.Open(output, false);
        C.ChartSpace space = Assert.Single(ChartSheet(package).DrawingsPart!.ChartParts).ChartSpace;
        Assert.Contains(space.Descendants<C.Formula>(), formula => formula.Text == "'Native Data'!$B$2:$B$3");
        Assert.Equal(cache, space.Descendants<C.NumberingCache>().Any());
        // Header-linked series labels retain their small native text cache independently of value caching.
        Assert.NotEmpty(space.Descendants<C.SeriesText>().SelectMany(text => text.Descendants<C.StringCache>()));
        Assert.Empty(ChartSheet(package).TableDefinitionParts);
        Assert.Empty(new OpenXmlValidator().Validate(package));
    }

    [Fact]
    public void DescriptorKeepsHorizontalCrossSheetLinksAndExplicitComboCache() {
        using var workbook = ExcelDocument.Create();
        ExcelSheet source = workbook.AddSheet("Source Data");
        ExcelSheet target = workbook.AddSheet("Summary");
        var data = new ExcelChartData(new[] { "Jan", "Feb" }, new[] {
            new ExcelChartSeries("Revenue", new[] { 10d, 20d }, OfficeChartType.Column),
            new ExcelChartSeries("Margin", new[] { 2d, 3d }, OfficeChartType.Line, OfficeChartAxisGroup.Secondary)
        });
        ExcelChartDataRange range = source.WriteChartData(data, 2, 3, orientation: ExcelChartDataOrientation.Horizontal);
        target.AddChart(OfficeChartType.Column, range, new ExcelChartOptions { CachedData = data, Row = 3, Column = 2 });
        using var output = Save(workbook);
        using var package = SpreadsheetDocument.Open(output, false);
        C.ChartSpace space = Assert.Single(ChartSheet(package).DrawingsPart!.ChartParts).ChartSpace;
        Assert.Single(space.Descendants<C.BarChart>());
        Assert.Single(space.Descendants<C.LineChart>());
        Assert.Equal(2, space.Descendants<C.ValueAxis>().Count());
        Assert.Contains(space.Descendants<C.Formula>(), formula => formula.Text == "'Source Data'!$D$3:$E$3");
        Assert.NotEmpty(space.Descendants<C.NumberingCache>());
        Assert.Empty(new OpenXmlValidator().Validate(package));
    }

    [Theory]
    [InlineData(true, true)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(false, false)]
    public void TableSourceUsesNativeHeadersAndExcludesTotals(bool headers, bool cache) {
        using var workbook = ExcelDocument.Create();
        ExcelSheet sheet = workbook.AddSheet("Data");
        using DataTable rows = SampleTable();
        ExcelTable table = sheet.AddTable(rows, new ExcelTableOptions { Name = "RevenueTable", IncludeHeaders = headers });
        Assert.Same(sheet, table.Sheet);
        Assert.Equal(headers ? "A1:B3" : "A1:B2", table.Range);
        // Native totals metadata denotes the final row in the table's current reference.
        table.Resize(headers ? "A1:B4" : "A1:B3").SetTotals(new Dictionary<string, ExcelTableTotalsFunction> {
            ["Revenue"] = ExcelTableTotalsFunction.Sum
        });
        sheet.CellValue(headers ? 4 : 3, 1, "Total");
        sheet.CellValue(headers ? 4 : 3, 2, 30d);
        sheet.AddChart(OfficeChartType.Column, "RevenueTable", new ExcelChartOptions { IncludeCachedData = cache });
        using var output = Save(workbook);
        using var package = SpreadsheetDocument.Open(output, false);
        C.ChartSpace space = Assert.Single(ChartSheet(package).DrawingsPart!.ChartParts).ChartSpace;
        string expected = headers ? "'Data'!$B$2:$B$3" : "'Data'!$B$1:$B$2";
        Assert.Contains(space.Descendants<C.Formula>(), formula => formula.Text == expected);
        Assert.Equal(cache, space.Descendants<C.NumberingCache>().Any());
        Assert.Equal("Revenue", Assert.Single(Assert.Single(space.Descendants<C.SeriesText>()).Descendants<C.NumericValue>()).Text);
        Assert.Empty(new OpenXmlValidator().Validate(package));
    }

    [Fact]
    public void NativeSeriesRetainsScatterCoordinatesAndExplicitSeriesRangeApis() {
        using var workbook = ExcelDocument.Create();
        ExcelSheet sheet = workbook.AddSheet("Data");
        var scatter = new ExcelChartData(Array.Empty<string>(), new[] {
            new ExcelChartSeries("Points", new[] { 20d, 40d }, new[] { 1d, 3d }, OfficeChartType.Scatter)
        });
        sheet.AddChart(OfficeChartType.Scatter, scatter);
        sheet.CellValues(new[] { (1, 1, (object)1d), (2, 1, (object)3d),
            (1, 2, (object)20d), (2, 2, (object)40d), (1, 3, (object)4d), (2, 3, (object)8d) });
        sheet.AddScatterChartFromRanges(new[] { new ExcelChartSeriesRange("XY", "A1:A2", "B1:B2") }, 20, 1);
        sheet.AddBubbleChartFromRanges(new[] { new ExcelChartSeriesRange("Sized", "A1:A2", "B1:B2", "C1:C2") }, 40, 1);
        using var output = Save(workbook);
        using var package = SpreadsheetDocument.Open(output, false);
        ChartPart[] parts = ChartSheet(package).DrawingsPart!.ChartParts.ToArray();
        Assert.Equal(3, parts.Length);
        Assert.Equal(2, parts.Count(part => part.ChartSpace.Descendants<C.ScatterChart>().Any()));
        C.BubbleChart bubble = Assert.Single(parts.SelectMany(part => part.ChartSpace.Descendants<C.BubbleChart>()));
        Assert.Contains(bubble.Descendants<C.Formula>(), formula => formula.Text == "'Data'!C1:C2");
        Assert.Empty(new OpenXmlValidator().Validate(package));
    }

    [Fact]
    public void TableAuthoringReturnsAttachedTablesAndSelectorProjectionEnumeratesOnce() {
        using var workbook = ExcelDocument.Create();
        ExcelSheet sheet = workbook.AddSheet("Tables");
        int enumerations = 0;
        IEnumerable<int> Rows() {
            if (++enumerations != 1) throw new InvalidOperationException("Second enumeration");
            yield return 1;
            yield return 2;
        }
        ExcelTable selected = sheet.AddTable(Rows(), new (string Header, Func<int, object?> Selector)[] {
            ("Code", value => value), ("Label", value => value == 1 ? null : string.Empty)
        }, new ExcelTableOptions { Name = "Selected" }, row: 3, column: 2);
        Assert.Equal(1, enumerations);
        Assert.Same(sheet, selected.Sheet);
        Assert.Equal("B3:C5", selected.Range);
        ExcelTable typed = sheet.AddTable(new[] { new { Region = "North", Revenue = 7m } },
            new ExcelTableOptions { Name = "Typed" }, row: 10, column: 2);
        using DataTable rows = SampleTable();
        ExcelTable native = sheet.AddTable(rows, new ExcelTableOptions { Name = "Native", IncludeAutoFilter = false }, row: 15, column: 2);
        Assert.Same(sheet, typed.Sheet);
        Assert.Same(sheet, native.Sheet);
        ExcelTable empty = sheet.AddTable(Array.Empty<int>(),
            new (string Header, Func<int, object?> Selector)[] { ("Code", value => value) },
            new ExcelTableOptions { Name = "Empty" }, row: 20, column: 2);
        Assert.Equal("B20:B20", empty.Range);
        using var output = Save(workbook);
        using var package = SpreadsheetDocument.Open(output, false);
        S.Table[] tables = package.WorkbookPart!.WorksheetParts.SelectMany(part => part.TableDefinitionParts).Select(part => part.Table!).ToArray();
        Assert.Equal(4, tables.Length);
        Assert.Null(tables.Single(table => table.Name!.Value == "Native").AutoFilter);
        Assert.Empty(new OpenXmlValidator().Validate(package));
    }

    [Fact]
    public void PlainDataKeepsUnstructuredCellsAndUnrestrictedSelectorHeaders() {
        using var workbook = ExcelDocument.Create();
        ExcelSheet sheet = workbook.AddSheet("Plain");
        using DataTable rows = SampleTable();
        sheet.WriteData(rows, row: 2, column: 3, includeHeaders: false);
        sheet.WriteData(new[] { new { Region = "West", Revenue = 9m } }, includeHeaders: false, row: 6);
        sheet.WriteData(new[] { 1, 2 }, includeHeaders: true, row: 10,
            ("", value => value == 1 ? null : DBNull.Value),
            ("Same", value => string.Empty), ("Same", value => value));
        Assert.Empty(sheet.GetTables());
        using var output = Save(workbook);
        using var package = SpreadsheetDocument.Open(output, false);
        WorksheetPart part = Assert.Single(package.WorkbookPart!.WorksheetParts);
        Assert.Empty(part.TableDefinitionParts);
        Assert.Null(part.Worksheet.GetFirstChild<S.AutoFilter>());
        Dictionary<string, S.Cell> cells = part.Worksheet.Descendants<S.Cell>().ToDictionary(cell => cell.CellReference!.Value!);
        Assert.Contains("C2", cells.Keys);
        Assert.Contains("D3", cells.Keys);
        foreach (string reference in new[] { "A11", "A12" }) {
            Assert.Null(cells[reference].CellValue);
            Assert.Null(cells[reference].InlineString);
            Assert.Null(cells[reference].DataType);
        }
        Assert.Contains("B11", cells.Keys);
        Assert.Contains("B12", cells.Keys);
        Assert.Empty(new OpenXmlValidator().Validate(package));
    }

    [Fact]
    public void TableAndPlainDataCancellationRetainTheSuppliedToken() {
        using var workbook = ExcelDocument.Create();
        ExcelSheet sheet = workbook.AddSheet("Data");
        using DataTable rows = SampleTable();
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.Equal(cancellation.Token, Assert.Throws<OperationCanceledException>(() =>
            sheet.WriteData(rows, cancellationToken: cancellation.Token)).CancellationToken);
        Assert.Equal(cancellation.Token, Assert.Throws<OperationCanceledException>(() =>
            sheet.AddTable(rows, new ExcelTableOptions { CancellationToken = cancellation.Token })).CancellationToken);
        Assert.Empty(sheet.GetTables());
    }

    private static ExcelChartData SampleData() => new(new[] { "Jan", "Feb" },
        new[] { new ExcelChartSeries("Revenue", new[] { 10d, 20d }) });

    private static DataTable SampleTable() {
        var data = new DataTable("Sales");
        data.Columns.Add("Month", typeof(string));
        data.Columns.Add("Revenue", typeof(double));
        data.Rows.Add("Jan", 10d);
        data.Rows.Add("Feb", 20d);
        return data;
    }

    private static MemoryStream Save(ExcelDocument workbook) {
        var output = new MemoryStream();
        try {
            workbook.Save(output);
            output.Position = 0;
            return output;
        } catch {
            output.Dispose();
            throw;
        }
    }

    private static WorksheetPart ChartSheet(SpreadsheetDocument package) =>
        Assert.Single(package.WorkbookPart!.WorksheetParts, part => part.DrawingsPart?.ChartParts.Any() == true);
}
