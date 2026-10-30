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

namespace OfficeIMO.Tests;

public sealed class ExcelNativeAuthoringReviewRegressions {
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void EffectiveSelectorSchemaFitsLastColumnWithoutReadingFirstExcludedSelector(bool useInclude) {
        using var workbook = ExcelDocument.Create();
        ExcelSheet sheet = workbook.AddSheet("Data");
        var options = new ExcelTableOptions { Name = "Selected" };
        if (useInclude) options.Include.Add("Keep");
        else options.Exclude.Add("Ignored");
        ExcelTable table = sheet.AddTable(new[] { 7 }, new (string, Func<int, object?>)[] {
            ("Ignored", _ => throw new InvalidOperationException("Excluded selector ran")),
            ("Keep", value => value)
        }, options, column: 16384);
        Assert.Equal("XFD1:XFD2", table.Range);
        Assert.Equal(7, sheet.CellAt(2, 16384).GetValue<int>());
        using var saved = new MemoryStream();
        workbook.Save(saved);
        saved.Position = 0;
        using var package = SpreadsheetDocument.Open(saved, false);
        S.Table native = Assert.Single(Assert.Single(package.WorkbookPart!.WorksheetParts).TableDefinitionParts).Table!;
        Assert.Equal("Keep", Assert.Single(native.TableColumns!.Elements<S.TableColumn>()).Name!.Value);
        Assert.Empty(new OpenXmlValidator().Validate(package));
    }

    [Fact]
    public void SelectorProjectionOrdersAndFormatsOriginalMembersExactlyOnce() {
        using var workbook = ExcelDocument.Create();
        ExcelSheet sheet = workbook.AddSheet("Data");
        int formatted = 0, transformed = 0, selected = 0;
        var options = new ExcelTableOptions {
            Include = { "Code", "Amount" }, Order = { "Amount", "Code" },
            HeaderRenames = { ["Amount"] = "Revenue" },
            HeaderTransform = name => { transformed++; return "Display " + name; }
        };
        options.Formatters["Amount"] = value => { formatted++; return "USD " + value; };
        sheet.AddTable(new[] { 2, 3 }, new (string, Func<int, object?>)[] {
            ("Ignored", _ => throw new InvalidOperationException("Excluded selector ran")),
            ("Code", value => { selected++; return value; }), ("Amount", value => value * 10)
        }, options);
        Assert.Equal(2, selected);
        Assert.Equal(2, formatted);
        Assert.Equal(1, transformed);
        Assert.Equal("Revenue", sheet.CellAt(1, 1).GetValue<string>());
        Assert.Equal("Display Code", sheet.CellAt(1, 2).GetValue<string>());
        Assert.Equal("USD 20", sheet.CellAt(2, 1).GetValue<string>());
        Assert.Equal("USD 30", sheet.CellAt(3, 1).GetValue<string>());
        Assert.Equal(3, sheet.CellAt(3, 2).GetValue<int>());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SelectorCancellationStopsSourceReadsAndDisposesBeforeEvaluatingSelectors(bool cancelAtEnd) {
        using var workbook = ExcelDocument.Create();
        ExcelSheet sheet = workbook.AddSheet("Data");
        using var cancellation = new CancellationTokenSource();
        int reads = 0, selectors = 0;
        bool disposed = false;
        IEnumerable<int> Rows() {
            try {
                reads++;
                if (!cancelAtEnd) cancellation.Cancel();
                yield return 1;
                reads++;
                if (!cancelAtEnd) throw new InvalidOperationException("Read after cancellation");
                cancellation.Cancel();
            } finally {
                disposed = true;
            }
        }
        OperationCanceledException error = Assert.Throws<OperationCanceledException>(() =>
            sheet.AddTable(Rows(), new (string, Func<int, object?>)[] {
                ("Value", value => { selectors++; return value; })
            }, new ExcelTableOptions { CancellationToken = cancellation.Token }));
        Assert.Equal(cancellation.Token, error.CancellationToken);
        Assert.Equal(cancelAtEnd ? 2 : 1, reads);
        Assert.True(disposed);
        Assert.Equal(0, selectors);
        Assert.Empty(sheet.GetTables());
        using var saved = new MemoryStream();
        workbook.Save(saved);
        saved.Position = 0;
        using var package = SpreadsheetDocument.Open(saved, false);
        Assert.Empty(Assert.Single(package.WorkbookPart!.WorksheetParts).Worksheet.Descendants<S.Cell>());
    }

    [Theory]
    [InlineData("data", false)]
    [InlineData("data", true)]
    [InlineData("typed", false)]
    [InlineData("typed", true)]
    [InlineData("selectors", false)]
    [InlineData("selectors", true)]
    public void NativeTableAutoFitCanBeDisabledWithoutChangingTheDefault(string route, bool autoFit) {
        using var workbook = ExcelDocument.Create();
        ExcelSheet sheet = workbook.AddSheet("Data");
        var options = new ExcelTableOptions();
        Assert.True(options.AutoFit);
        options.AutoFit = autoFit;
        const string text = "An extended description requiring a much wider native column";
        if (route == "data") {
            using var rows = new DataTable();
            rows.Columns.Add("Description", typeof(string));
            rows.Rows.Add(text);
            sheet.AddTable(rows, options);
        } else if (route == "typed") {
            sheet.AddTable(new[] { new { Description = text } }, options);
        } else {
            sheet.AddTable(new[] { text }, new (string, Func<string, object?>)[] { ("Description", value => value) }, options);
        }
        using var saved = new MemoryStream();
        workbook.Save(saved);
        saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        ExcelSheet result = reopened.Sheets[0];
        Assert.Equal(text, result.CellAt(2, 1).GetValue<string>());
        if (autoFit) Assert.True(Assert.Single(result.GetColumnDefinitions()).Width?.CharacterCount > 20);
        else Assert.Empty(result.GetColumnDefinitions());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HeaderlessAppendPreservesPositionalCallerAndNewNamedTableSemantics(bool positional) {
        using var workbook = ExcelDocument.Create();
        ExcelSheet sheet = workbook.AddSheet("Data");
        using var rows = Sales();
        sheet.AddTable(rows, new ExcelTableOptions {
            Name = "Sales", IncludeHeaders = false, AutoFit = false,
            UseDefaultHeaderlessColumnNames = positional
        });
        using var appended = new DataTable();
        appended.Columns.Add(positional ? "Column2" : "Revenue", typeof(string));
        appended.Columns.Add(positional ? "Column1" : "Month", typeof(string));
        appended.Rows.Add("30", "Mar");
        Assert.Equal("A1:B3", sheet.AppendDataTableToTable(appended, "Sales"));
        using var saved = new MemoryStream();
        workbook.Save(saved);
        saved.Position = 0;
        using var inspection = new MemoryStream(saved.ToArray());
        using (var package = SpreadsheetDocument.Open(inspection, false)) {
            var part = Assert.Single(package.WorkbookPart!.WorksheetParts);
            Assert.Equal(positional ? new[] { "Column1", "Column2" } : new[] { "Month", "Revenue" },
                Assert.Single(part.TableDefinitionParts).Table!.TableColumns!.Elements<S.TableColumn>().Select(field => field.Name!.Value!));
            Assert.Empty(part.Worksheet.Elements<S.Columns>());
        }
        saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        Assert.Equal(positional ? "30" : "Mar", reopened.Sheets[0].CellAt(3, 1).GetValue<string>());
        Assert.Equal(positional ? "Mar" : "30", reopened.Sheets[0].CellAt(3, 2).GetValue<string>());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HeaderlessNativeChartRetainsEverySeriesNameAndTotalsOffsetsWithOptionalCaches(bool cache) {
        using var workbook = ExcelDocument.Create();
        ExcelSheet sheet = workbook.AddSheet("Native Data");
        using var rows = Sales();
        rows.Columns.Add("Cost", typeof(double));
        rows.Rows[0]["Cost"] = 4d;
        rows.Rows[1]["Cost"] = 8d;
        ExcelTable table = sheet.AddTable(rows, new ExcelTableOptions { Name = "Sales", IncludeHeaders = false }, row: 4, column: 3);
        table.Resize("C4:E6").SetTotals(new Dictionary<string, ExcelTableTotalsFunction> {
            ["Revenue"] = ExcelTableTotalsFunction.Sum, ["Cost"] = ExcelTableTotalsFunction.Sum
        });
        var authored = sheet.AddChart(OfficeChartType.Column3DClustered, "Sales", new ExcelChartOptions { IncludeCachedData = cache });
        Assert.True(authored.TryGetData(out var currentData));
        Assert.Equal(new[] { "Revenue", "Cost" }, currentData.Series.Select(series => series.Name));
        using var saved = new MemoryStream();
        workbook.Save(saved);
        saved.Position = 0;
        using var package = SpreadsheetDocument.Open(saved, false);
        C.ChartSpace chart = Assert.Single(Assert.Single(package.WorkbookPart!.WorksheetParts).DrawingsPart!.ChartParts).ChartSpace;
        Assert.Equal(new[] { "Revenue", "Cost" }, chart.Descendants<C.SeriesText>().Select(text => text.InnerText));
        Assert.Contains(chart.Descendants<C.Formula>(), formula => formula.Text == "'Native Data'!$C$4:$C$5");
        Assert.Contains(chart.Descendants<C.Formula>(), formula => formula.Text == "'Native Data'!$D$4:$D$5");
        Assert.Contains(chart.Descendants<C.Formula>(), formula => formula.Text == "'Native Data'!$E$4:$E$5");
        Assert.Equal(cache, chart.Descendants<C.NumberingCache>().Any());
        Assert.Empty(new OpenXmlValidator().Validate(package));
        using var source = new MemoryStream(saved.ToArray());
        using var reopened = ExcelDocument.Load(source);
        Assert.True(reopened.Sheets.Single().Charts.Single().TryGetData(out var savedData));
        Assert.Equal(new[] { "Revenue", "Cost" }, savedData.Series.Select(series => series.Name));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LinkedHeaderNamesUseCurrentWorksheetValues(bool cache) {
        using var workbook = ExcelDocument.Create();
        var sheet = workbook.AddSheet("Data");
        using var rows = Sales();
        sheet.WriteData(rows);
        var chart = sheet.AddChart(OfficeChartType.Column, "A1:B3", new ExcelChartOptions { IncludeCachedData = cache });
        sheet.CellValue(1, 2, "Updated revenue");
        Assert.True(chart.TryGetData(out var current));
        Assert.Equal("Updated revenue", Assert.Single(current.Series).Name);
    }

    [Fact]
    public void FinalSelectorCancellationStopsBeforeFormattingOrTableInsertion() {
        using var workbook = ExcelDocument.Create();
        var sheet = workbook.AddSheet("Data");
        using var cancellation = new CancellationTokenSource();
        int formats = 0;
        var options = new ExcelTableOptions { CancellationToken = cancellation.Token };
        options.Formatters["Value"] = value => { formats++; return value?.ToString() ?? string.Empty; };
        var columns = new (string Header, Func<int, object?> Selector)[] {
            ("Value", value => { cancellation.Cancel(); return value; })
        };
        var error = Assert.Throws<OperationCanceledException>(() => sheet.AddTable(new[] { 1 }, columns, options));
        Assert.Equal(cancellation.Token, error.CancellationToken);
        Assert.Equal(0, formats);
        Assert.Empty(sheet.GetTables());
    }

    private static DataTable Sales() {
        var rows = new DataTable();
        rows.Columns.Add("Month", typeof(string));
        rows.Columns.Add("Revenue", typeof(double));
        rows.Rows.Add("Jan", 10d);
        rows.Rows.Add("Feb", 20d);
        return rows;
    }
}
