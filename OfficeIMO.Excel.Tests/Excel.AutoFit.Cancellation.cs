using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelAutoFitCancellationTests {
    [Theory]
    [InlineData("all", false)]
    [InlineData("selected", false)]
    [InlineData("except", false)]
    [InlineData("all", true)]
    [InlineData("selected", true)]
    [InlineData("except", true)]
    public void PreCanceledAutoFitPreservesDeferredValuesAndWidths(string route, bool alreadyFit) {
        using var source = new MemoryStream();
        using var document = ExcelDocument.Create(source);
        ExcelSheet sheet = document.AddSheet("Data");
        using DataTable table = CreateTable();
        sheet.InsertDataTableAsTable(table);
        if (alreadyFit) sheet.AutoFitColumns();
        // Inspect a separate workbook: inspecting this sheet would materialize its deferred source.
        var widths = ExpectedTableWidths(alreadyFit);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        OperationCanceledException error = Assert.Throws<OperationCanceledException>(() =>
            FitColumns(sheet, route, cancellation.Token));

        Assert.Equal(cancellation.Token, error.CancellationToken);
        Assert.Equal("An extended description that needs a wider column", table.Rows[0][0]);
        Assert.Equal(42D, table.Rows[0][1]);
        AssertSavedValuesAndWidths(document, widths);
    }

    [Theory]
    [InlineData("all")]
    [InlineData("selected")]
    [InlineData("except")]
    public void PreCanceledAutoFitObservesCancellationWhenThereAreNoColumnsToFit(string route) {
        using var document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddSheet("Data");
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() => {
            if (route == "all") sheet.AutoFitColumns(ct: cancellation.Token);
            else if (route == "selected") sheet.AutoFitColumnsFor(Array.Empty<int>(), ct: cancellation.Token);
            else sheet.AutoFitColumnsExcept(Array.Empty<int>(), ct: cancellation.Token);
        });

        Assert.Empty(sheet.GetColumnDefinitions());
        using var saved = new MemoryStream();
        document.Save(saved);
        using var reopened = ExcelDocument.Load(saved);
        Assert.Empty(reopened.Sheets[0].GetColumnDefinitions());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CancellationWhileEnumeratingColumnsDoesNotBecomeNoWorkSuccess(bool except) {
        using var source = new MemoryStream();
        using var document = ExcelDocument.Create(source);
        ExcelSheet sheet = document.AddSheet("Data");
        using DataTable table = CreateTable();
        sheet.InsertDataTableAsTable(table);
        var widths = Array.Empty<(int Start, int End, double? Width, bool Custom)>();
        using var cancellation = new CancellationTokenSource();
        IEnumerable<int> Columns() {
            cancellation.Cancel();
            if (except) {
                yield return 1;
                yield return 2;
            } else {
                yield return 0;
            }
        }

        Assert.Throws<OperationCanceledException>(() => {
            if (except) sheet.AutoFitColumnsExcept(Columns(), ct: cancellation.Token);
            else sheet.AutoFitColumnsFor(Columns(), ct: cancellation.Token);
        });

        AssertSavedValuesAndWidths(document, widths);
    }

    [Theory]
    [InlineData("all", false)]
    [InlineData("selected", false)]
    [InlineData("selectedAll", false)]
    [InlineData("all", true)]
    [InlineData("selected", true)]
    [InlineData("selectedAll", true)]
    public void CancellationDuringDirectMeasurementPreservesTheCandidateForLaterSave(string route, bool hasLaterRow) {
        using var source = new MemoryStream();
        using var document = ExcelDocument.Create(source);
        document.DateTimeOffsetWriteStrategy = value => value.UtcDateTime;
        ExcelSheet sheet = document.AddSheet("Data");
        using var table = new DataTable("Data");
        table.Columns.Add("Description", typeof(string));
        table.Columns.Add("When", typeof(DateTimeOffset));
        var date = new DateTimeOffset(2026, 10, 11, 8, 30, 0, TimeSpan.Zero);
        table.Rows.Add("Retained source", date);
        if (hasLaterRow) table.Rows.Add("Retained later row", date.AddDays(1));
        sheet.InsertDataTableAsTable(table);
        var widths = Array.Empty<(int Start, int End, double? Width, bool Custom)>();
        using var cancellation = new CancellationTokenSource();
        bool measuredDate = false;
        document.DateTimeOffsetWriteStrategy = value => {
            measuredDate = true;
            cancellation.Cancel();
            return value.UtcDateTime;
        };

        OperationCanceledException error = Assert.Throws<OperationCanceledException>(() =>
            FitColumns(sheet, route, cancellation.Token));

        Assert.True(measuredDate);
        Assert.Equal(cancellation.Token, error.CancellationToken);
        Assert.Equal(date, table.Rows[0][1]);
        document.DateTimeOffsetWriteStrategy = value => value.UtcDateTime;
        using var saved = new MemoryStream();
        document.Save(saved);
        Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
        using var reopened = ExcelDocument.Load(saved);
        Assert.Equal(widths, ColumnState(reopened.Sheets[0]));
        Assert.Equal("Retained source", reopened.Sheets[0].CellAt(2, 1).GetValue<string>());
        Assert.Equal(date.UtcDateTime.ToOADate(), reopened.Sheets[0].CellAt(2, 2).GetValue<double>());
        if (hasLaterRow) {
            Assert.Equal("Retained later row", reopened.Sheets[0].CellAt(3, 1).GetValue<string>());
            Assert.Equal(date.AddDays(1).UtcDateTime.ToOADate(), reopened.Sheets[0].CellAt(3, 2).GetValue<double>());
        }
    }

    [Fact]
    public void SequentialAutoFitDoesNotApplyWidthsWhenMeasurementCompletionCancels() {
        using var document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddSheet("Data");
        sheet.CellValue(1, 1, "A long ordinary cell value requiring a wider column");
        var widths = ColumnState(sheet);
        using var cancellation = new CancellationTokenSource();
        document.Execution.OnTiming = (operation, _) => {
            if (operation == "AutoFitColumns.CalculateWidths") cancellation.Cancel();
        };

        Assert.Throws<OperationCanceledException>(() => sheet.AutoFitColumns(ExcelExecutionMode.Sequential, cancellation.Token));

        Assert.Equal(widths, ColumnState(sheet));
        document.Execution.OnTiming = null;
        using var saved = new MemoryStream();
        document.Save(saved);
        using var reopened = ExcelDocument.Load(saved);
        Assert.Equal(widths, ColumnState(reopened.Sheets[0]));
        Assert.Equal("A long ordinary cell value requiring a wider column", reopened.Sheets[0].CellAt(1, 1).GetValue<string>());
    }

    private static void FitColumns(ExcelSheet sheet, string route, CancellationToken cancellationToken) {
        if (route == "all") sheet.AutoFitColumns(ct: cancellationToken);
        else if (route == "selected") sheet.AutoFitColumnsFor(new[] { 2 }, ct: cancellationToken);
        else if (route == "selectedAll") sheet.AutoFitColumnsFor(new[] { 1, 2 }, ct: cancellationToken);
        else sheet.AutoFitColumnsExcept(new[] { 1 }, ct: cancellationToken);
    }

    private static DataTable CreateTable() {
        var table = new DataTable("Data");
        table.Columns.Add("Description", typeof(string));
        table.Columns.Add("Amount", typeof(double));
        table.Rows.Add("An extended description that needs a wider column", 42D);
        return table;
    }

    private static (int Start, int End, double? Width, bool Custom)[] ColumnState(ExcelSheet sheet) =>
        sheet.GetColumnDefinitions().Select(column =>
            (column.StartIndex, column.EndIndex, (column.Width?.CharacterCount), column.CustomWidth)).ToArray();

    private static (int Start, int End, double? Width, bool Custom)[] ExpectedTableWidths(bool alreadyFit) {
        using var source = new MemoryStream();
        using var document = ExcelDocument.Create(source);
        ExcelSheet sheet = document.AddSheet("Data");
        using DataTable table = CreateTable();
        sheet.InsertDataTableAsTable(table);
        if (alreadyFit) sheet.AutoFitColumns();
        using var saved = new MemoryStream();
        document.Save(saved);
        using var reopened = ExcelDocument.Load(saved);
        return ColumnState(reopened.Sheets[0]);
    }

    private static void AssertSavedValuesAndWidths(ExcelDocument document,
        (int Start, int End, double? Width, bool Custom)[] widths) {
        using var saved = new MemoryStream();
        document.Save(saved);
        Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
        using var reopened = ExcelDocument.Load(saved);
        Assert.Equal(widths, ColumnState(reopened.Sheets[0]));
        Assert.Equal("An extended description that needs a wider column", reopened.Sheets[0].CellAt(2, 1).GetValue<string>());
        Assert.Equal(42D, reopened.Sheets[0].CellAt(2, 2).GetValue<double>());
    }
}
