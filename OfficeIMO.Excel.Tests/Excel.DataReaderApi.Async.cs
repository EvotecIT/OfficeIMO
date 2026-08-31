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
}
#endif
