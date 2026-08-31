#if NET8_0_OR_GREATER
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Fact]
    public async Task OpenDataReaderAsyncCursorTraversesRowsAndResultSets() {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.AsyncReader.{Guid.NewGuid():N}.xlsx");
        try {
            using (ExcelDocument document = ExcelDocument.Create(path)) {
                ExcelSheet first = document.AddWorksheet("First");
                first.CellValue(1, 1, "Value");
                first.CellValue(2, 1, "Alpha");
                ExcelSheet second = document.AddWorksheet("Second");
                second.CellValue(1, 1, "Value");
                second.CellValue(2, 1, "Beta");
                document.Save();
            }

            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(path);
            Assert.True(await reader.ReadAsync(CancellationToken.None));
            Assert.Equal("Alpha", reader.GetString(0));
            Assert.False(await reader.ReadAsync(CancellationToken.None));
            Assert.True(await reader.NextResultAsync(CancellationToken.None));
            Assert.True(await reader.ReadAsync(CancellationToken.None));
            Assert.Equal("Beta", reader.GetString(0));
            Assert.False(await reader.NextResultAsync(CancellationToken.None));

            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();
            await Assert.ThrowsAsync<OperationCanceledException>(
                () => reader.ReadAsync(cancellation.Token));
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public async Task NextResultAsyncThreadsThePerCallTokenIntoSheetOpening() {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.AsyncNextResult.{Guid.NewGuid():N}.xlsx");
        using var cancellation = new CancellationTokenSource();
        bool cancelDuringNextSheet = false;
        try {
            using (ExcelDocument document = ExcelDocument.Create(path)) {
                ExcelSheet first = document.AddWorksheet("First");
                first.CellValue(1, 1, "Value");
                first.CellValue(2, 1, "Alpha");
                ExcelSheet second = document.AddWorksheet("Second");
                second.CellValue(1, 1, "Value");
                for (int row = 2; row <= 512; row++) {
                    second.CellValue(row, 1, row == 2 ? "CancelHere" : "Value" + row);
                }
                document.Save();
            }

            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(
                path,
                new ExcelReadOptions {
                    InferSchema = true,
                    SchemaSampleRows = 512,
                    CellValueConverter = _ => {
                        if (Volatile.Read(ref cancelDuringNextSheet)) cancellation.Cancel();
                        return ExcelCellValue.NotHandled;
                    }
                });

            Volatile.Write(ref cancelDuringNextSheet, true);
            await Assert.ThrowsAsync<OperationCanceledException>(
                () => reader.NextResultAsync(cancellation.Token));
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public async Task NextResultAsyncDoesNotRetainCompletedPerCallTokenForLaterReads() {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.AsyncNextResultLifetime.{Guid.NewGuid():N}.xlsx");
        using var perCallCancellation = new CancellationTokenSource();
        try {
            using (ExcelDocument document = ExcelDocument.Create(path)) {
                ExcelSheet first = document.AddWorksheet("First");
                first.CellValue(1, 1, "Value");
                first.CellValue(2, 1, "Alpha");
                ExcelSheet second = document.AddWorksheet("Second");
                second.CellValue(1, 1, "Value");
                second.CellValue(2, 1, "Beta");
                document.Save();
            }

            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(path);
            Assert.True(await reader.NextResultAsync(perCallCancellation.Token));
            perCallCancellation.Cancel();

            Assert.True(await reader.ReadAsync(CancellationToken.None));
            Assert.Equal("Beta", reader.GetString(0));
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData(ExcelFileFormat.Xlsb)]
    public async Task ReadAsyncThreadsThePerCallTokenIntoBinaryWorksheetReads(ExcelFileFormat format) {
        byte[] workbook;
        using (ExcelDocument document = ExcelDocument.Create()) {
            ExcelSheet sheet = document.AddWorksheet("Data");
            for (int column = 1; column <= 256; column++) {
                sheet.CellValue(1, column, "Column" + column);
                sheet.CellValue(2, column, column);
            }
            workbook = document.ToBytes(format);
        }

        using var cancellation = new CancellationTokenSource();
        bool cancelDuringRead = false;
        using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(
            workbook,
            new ExcelReadOptions {
                InferSchema = false,
                CellValueConverter = _ => {
                    if (Volatile.Read(ref cancelDuringRead)) cancellation.Cancel();
                    return ExcelCellValue.NotHandled;
                }
            });

        Volatile.Write(ref cancelDuringRead, true);
        Task<bool> canceledRead = reader.ReadAsync(cancellation.Token);
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => canceledRead);
        Assert.True(canceledRead.IsCanceled);
    }

    [Theory]
    [InlineData(ExcelFileFormat.Xlsx)]
    [InlineData(ExcelFileFormat.Xlsb)]
    [InlineData(ExcelFileFormat.Xls)]
    public async Task ReadAsyncCompletesRowAdvanceWithoutThreadPoolDispatch(ExcelFileFormat format) {
        byte[] workbook;
        using (ExcelDocument document = ExcelDocument.Create()) {
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Value");
            sheet.CellValue(2, 1, "Alpha");
            workbook = document.ToBytes(format);
        }

        int callerThread = Environment.CurrentManagedThreadId;
        int converterThread = 0;
        using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(
            workbook,
            new ExcelReadOptions {
                InferSchema = false,
                CellValueConverter = value => {
                    Volatile.Write(ref converterThread, Environment.CurrentManagedThreadId);
                    return ExcelCellValue.NotHandled;
                }
            });

        Task<bool> read = reader.ReadAsync(CancellationToken.None);

        Assert.True(read.IsCompleted);
        Assert.True(await read);
        Assert.Equal(callerThread, Volatile.Read(ref converterThread));
    }
}
#endif
