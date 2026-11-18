using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    public sealed class BufferedCancellationRow {
        public string? Name { get; set; }
    }

    /// <summary>
    /// A small range is read into memory before its first row is returned, so the token has to be checked when each
    /// buffered row is handed out, not only while the worksheet is being parsed.
    /// </summary>
    [Theory]
    [InlineData("ReadObjects")]
    [InlineData("ReadObjectsOfT")]
    [InlineData("ReadRangeStream")]
    [InlineData("ReadRangeStreamOnlyChunk")]
    public void SheetReader_ObservesCancellationBetweenBufferedRows(string api) {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.BufferedCancellation.{Guid.NewGuid():N}.xlsx");
        try {
            using (var document = ExcelDocument.Create(path)) {
                ExcelSheet sheet = document.AddSheet("Data");
                sheet.Cell(1, 1).SetValue("Name");
                for (int row = 2; row <= 4; row++) sheet.Cell(row, 1).SetValue("Row " + row);
                document.Save();
            }

            using var cancellation = new CancellationTokenSource();
            using var reader = ExcelDocumentReader.Open(path);
            var data = reader.GetSheet(1);
            System.Collections.IEnumerable rows = api switch {
                "ReadObjects" => data.ReadObjects("A1:A4", ct: cancellation.Token),
                "ReadObjectsOfT" => data.ReadObjects<BufferedCancellationRow>("A1:A4", ct: cancellation.Token),
                "ReadRangeStream" => data.ReadRangeStream("A1:A4", chunkRows: 1, ct: cancellation.Token),
                // Cancelled after the only chunk: the next advance reports the cancel instead of completing.
                _ => data.ReadRangeStream("A1:A4", chunkRows: 100, ct: cancellation.Token)
            };

            System.Collections.IEnumerator enumerator = rows.GetEnumerator();
            Assert.True(enumerator.MoveNext());
            cancellation.Cancel();
            Assert.ThrowsAny<OperationCanceledException>(() => enumerator.MoveNext());
        } finally {
            File.Delete(path);
        }
    }
}
