using OfficeIMO.Excel.Xlsb;
using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Read;
using System.Threading;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Fact]
    public void XlsbTabularReader_DisposesItsRecordReaderWhenConstructionFails() {
        using var worksheetPart = new TrackingReadStream(new byte[] { 0x80 });

        Assert.Throws<EndOfStreamException>(() =>
            new XlsbTabularDataReader(
                worksheetPart,
                Array.Empty<string>(),
                Array.Empty<bool>(),
                uses1904DateSystem: false,
                hasHeaderRow: true,
                new ExcelReadOptions(),
                new XlsbImportOptions(),
                new XlsbRecordReadBudget(100),
                new XlsbCellReadBudget(100),
                CancellationToken.None));

        Assert.True(worksheetPart.WasDisposed);
    }

    [Fact]
    public void XlsbTabularReader_EmitsBlankRowsForPhysicalRowGaps() {
        using var worksheetPart = CreateTabularWorksheet(
            (0, 0U),
            (2, 1U));
        using var reader = CreateTabularReader(
            worksheetPart,
            new[] { "Value", "Data" },
            hasHeaderRow: true,
            new XlsbCellReadBudget(10));

        Assert.True(reader.Read());
        Assert.True(reader.IsDBNull(0));
        Assert.True(reader.Read());
        Assert.Equal("Data", reader.GetString(0));
        Assert.False(reader.Read());
    }

    [Fact]
    public void XlsbTabularReader_SharesCellBudgetAcrossWorksheetReaders() {
        var cellBudget = new XlsbCellReadBudget(1);
        using (var firstWorksheet = CreateTabularWorksheet((0, 0U)))
        using (var firstReader = CreateTabularReader(
            firstWorksheet,
            new[] { "First" },
            hasHeaderRow: false,
            cellBudget)) {
            Assert.True(firstReader.Read());
            Assert.Equal("First", firstReader.GetString(0));
        }

        using var secondWorksheet = CreateTabularWorksheet((0, 0U));
        using var secondReader = CreateTabularReader(
            secondWorksheet,
            new[] { "Second" },
            hasHeaderRow: false,
            cellBudget);

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => secondReader.Read());
        Assert.Contains("workbook", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("populated cells", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ExcelDocumentReader_DisposesOpenedDocumentWhenInitializationFails() {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.Excel.ReaderDispose.{Guid.NewGuid():N}.xlsx");
        try {
            using (var document = ExcelDocument.Create(path)) {
                document.AddWorksheet("Data").CellValue(1, 1, "Value");
                document.Save();
            }

            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();
            Assert.Throws<OperationCanceledException>(() =>
                ExcelDocumentReader.Open(
                    path,
                    new ExcelReadOptions { CancellationToken = cancellation.Token }));

            using var exclusive = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            Assert.True(exclusive.CanRead);
        } finally {
            File.Delete(path);
        }
    }

    private static XlsbTabularDataReader CreateTabularReader(
        Stream worksheetPart,
        IReadOnlyList<string> sharedStrings,
        bool hasHeaderRow,
        XlsbCellReadBudget cellBudget) =>
        new(
            worksheetPart,
            sharedStrings,
            Array.Empty<bool>(),
            uses1904DateSystem: false,
            hasHeaderRow,
            new ExcelReadOptions(),
            new XlsbImportOptions(),
            new XlsbRecordReadBudget(100),
            cellBudget,
            CancellationToken.None);

    private static MemoryStream CreateTabularWorksheet(params (int RowIndex, uint SharedStringIndex)[] rows) {
        var stream = new MemoryStream();
        int lastRow = rows.Length == 0 ? 0 : rows.Max(static row => row.RowIndex);
        XlsbRecordWriter.Write(stream, 148, CreateWorksheetDimensionPayload(0, lastRow, 0, 0));
        XlsbRecordWriter.Write(stream, 145);
        foreach ((int rowIndex, uint sharedStringIndex) in rows) {
            XlsbRecordWriter.Write(stream, 0, CreateTabularRowHeaderPayload(rowIndex));
            XlsbRecordWriter.Write(stream, 7, CreateSharedStringCellPayload(0, sharedStringIndex));
        }

        XlsbRecordWriter.Write(stream, 146);
        stream.Position = 0;
        return stream;
    }

    private static byte[] CreateWorksheetDimensionPayload(
        int firstRow,
        int lastRow,
        int firstColumn,
        int lastColumn) {
        using var stream = new MemoryStream();
        using var writer = new BinaryWriter(stream);
        writer.Write(firstRow);
        writer.Write(lastRow);
        writer.Write(firstColumn);
        writer.Write(lastColumn);
        return stream.ToArray();
    }

    private static byte[] CreateTabularRowHeaderPayload(int rowIndex) {
        byte[] payload = new byte[17];
        using var stream = new MemoryStream(payload, writable: true);
        using var writer = new BinaryWriter(stream);
        writer.Write(rowIndex);
        return payload;
    }

    private static byte[] CreateSharedStringCellPayload(int column, uint sharedStringIndex) {
        using var stream = new MemoryStream();
        using var writer = new BinaryWriter(stream);
        writer.Write(column);
        writer.Write(0U);
        writer.Write(sharedStringIndex);
        return stream.ToArray();
    }

    private sealed class TrackingReadStream : MemoryStream {
        internal TrackingReadStream(byte[] bytes) : base(bytes, writable: false) {
        }

        internal bool WasDisposed { get; private set; }

        protected override void Dispose(bool disposing) {
            WasDisposed = true;
            base.Dispose(disposing);
        }
    }
}
