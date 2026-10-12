using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelDataReaderAutoFitCancellationTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CancellationDuringDeferredReaderAutoFitDoesNotPublishTheImport(bool hasLaterRow) {
        using var document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Data");
        using DataTable table = CreateDates(hasLaterRow);
        using var cancellation = new CancellationTokenSource();
        bool measuredDate = false;
        document.DateTimeOffsetWriteStrategy = value => {
            measuredDate = true;
            cancellation.Cancel();
            return value.UtcDateTime;
        };
        using var reader = table.CreateDataReader();

        OperationCanceledException error = Assert.Throws<OperationCanceledException>(() =>
            sheet.InsertDataReaderWithResult(reader, autoFit: true, ct: cancellation.Token));

        Assert.True(measuredDate);
        Assert.Equal(cancellation.Token, error.CancellationToken);
        Assert.Equal(Date(), table.Rows[0][0]);
        AssertEmptySavedSheet(document);

        document.DateTimeOffsetWriteStrategy = value => value.UtcDateTime;
        using var retryReader = table.CreateDataReader();
        ExcelDataReaderInsertResult result = sheet.InsertDataReaderWithResult(retryReader, autoFit: true);
        Assert.Equal(hasLaterRow ? "A1:A3" : "A1:A2", result.Range);
        using var saved = new MemoryStream();
        document.Save(saved);
        using var reopened = ExcelDocument.Load(saved);
        Assert.Single(reopened.GetTables());
        Assert.True(Assert.Single(reopened.Sheets[0].GetColumnDefinitions()).Width > 0D);
        Assert.Equal(Date().UtcDateTime.ToOADate(), reopened.Sheets[0].CellAt(2, 1).GetValue<double>());
        if (hasLaterRow) {
            Assert.Equal(Date().AddDays(1).UtcDateTime.ToOADate(), reopened.Sheets[0].CellAt(3, 1).GetValue<double>());
        }
    }

    [Fact]
    public void PreCanceledReaderAutoFitObservesCancellationForAnEmptyImport() {
        using var document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Data");
        using var table = new DataTable("Data");
        table.Columns.Add("When", typeof(DateTimeOffset));
        using var reader = table.CreateDataReader();
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        OperationCanceledException error = Assert.Throws<OperationCanceledException>(() =>
            sheet.InsertDataReaderWithResult(reader, includeHeaders: false, autoFit: true, ct: cancellation.Token));

        Assert.Equal(cancellation.Token, error.CancellationToken);
        AssertEmptySavedSheet(document);
    }

    [Fact]
    public void ReaderAutoFitOnANonemptySheetDoesNotApplyWidthsAfterCancellation() {
        using var document = ExcelDocument.Create();
        document.DateTimeOffsetWriteStrategy = value => value.UtcDateTime;
        ExcelSheet sheet = document.AddWorksheet("Data");
        sheet.CellValue(1, 1, "Existing value");
        var widths = sheet.GetColumnDefinitions().Select(column => column.Width).ToArray();
        using DataTable table = CreateDates(hasLaterRow: false);
        using var reader = table.CreateDataReader();
        using var cancellation = new CancellationTokenSource();
        document.Execution.OnTiming = (operation, _) => {
            if (operation == "AutoFitColumns.CalculateWidths") cancellation.Cancel();
        };

        OperationCanceledException error = Assert.Throws<OperationCanceledException>(() =>
            sheet.InsertDataReaderWithResult(reader, startRow: 2, autoFit: true, ct: cancellation.Token));

        Assert.Equal(cancellation.Token, error.CancellationToken);
        Assert.Equal(widths, sheet.GetColumnDefinitions().Select(column => column.Width));
        document.Execution.OnTiming = null;
        using var saved = new MemoryStream();
        document.Save(saved);
        using var reopened = ExcelDocument.Load(saved);
        Assert.Equal(widths, reopened.Sheets[0].GetColumnDefinitions().Select(column => column.Width));
        Assert.Equal("Existing value", reopened.Sheets[0].CellAt(1, 1).GetValue<string>());
        Assert.Equal("When", reopened.Sheets[0].CellAt(2, 1).GetValue<string>());
        Assert.Equal(Date().UtcDateTime.ToOADate(), reopened.Sheets[0].CellAt(3, 1).GetValue<double>());
    }

    private static DateTimeOffset Date() => new(2026, 10, 11, 8, 30, 0, TimeSpan.Zero);

    private static DataTable CreateDates(bool hasLaterRow) {
        var table = new DataTable("Data");
        table.Columns.Add("When", typeof(DateTimeOffset));
        table.Rows.Add(Date());
        if (hasLaterRow) table.Rows.Add(Date().AddDays(1));
        return table;
    }

    private static void AssertEmptySavedSheet(ExcelDocument document) {
        using var saved = new MemoryStream();
        document.Save(saved);
        using var reopened = ExcelDocument.Load(saved);
        Assert.Empty(reopened.GetTables());
        Assert.Empty(reopened.Sheets[0].GetColumnDefinitions());
        Assert.True(string.IsNullOrEmpty(reopened.Sheets[0].CellAt(1, 1).GetValue<string>()));
        Assert.True(string.IsNullOrEmpty(reopened.Sheets[0].CellAt(2, 1).GetValue<string>()));
    }
}
