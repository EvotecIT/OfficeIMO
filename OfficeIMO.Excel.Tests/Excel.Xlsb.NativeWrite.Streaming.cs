using OfficeIMO.Excel;
using OfficeIMO.Excel.Xlsb.Write;
using System.Threading;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void Xlsb_StreamingSnapshot_ReadsEachFieldOnceBeforeDestinationMutation(bool sharedStrings) {
            var rows = new SnapshotXlsbRows();
            DateTime expectedDate = (DateTime)rows.Values[0][2]!;
            using ExcelDocument document = ExcelDocument.Create();
            document.AddWorksheet("Data");
            var source = new ExcelDirectTabularSource("Data", rows, includeHeaders: true, preserveMissingValues: true);
            using var destination = new SnapshotXlsbStream(() => {
                Assert.Equal(rows.RowCount * rows.ColumnCount, rows.ValueReads);
                rows.Values[0][0] = "Mutated after capture";
                rows.Values[0][1] = -1;
                rows.Values[0][2] = DateTime.MinValue;
                rows.Values[1][3] = true;
            });

            Assert.True(XlsbNewPackageWriter.TryWriteDirectTabular(document, source, destination,
                CancellationToken.None, sharedStrings));
            Assert.Equal(rows.RowCount * rows.ColumnCount, rows.ValueReads);
            Assert.Equal(rows.ColumnCount, rows.HeaderReads);
            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(destination.ToArray());
            Assert.Equal("Text", reader.GetName(0));
            Assert.Equal("Number", reader.GetName(1));
            Assert.Equal("Date", reader.GetName(2));
            Assert.Equal("Boolean", reader.GetName(3));
            Assert.True(reader.Read());
            Assert.Equal("Zażółć 🚀", reader.GetString(0));
            Assert.Equal(42D, reader.GetDouble(1));
            Assert.Equal(expectedDate, reader.GetDateTime(2));
            Assert.True(reader.GetBoolean(3));
            Assert.True(reader.Read());
            Assert.Equal("", reader.GetString(0));
            Assert.True(reader.IsDBNull(1));
            Assert.True(reader.IsDBNull(2));
            Assert.False(reader.GetBoolean(3));
            Assert.True(reader.Read());
            for (int column = 0; column < 4; column++) Assert.True(reader.IsDBNull(column));
            Assert.False(reader.Read());
            Assert.False(reader.NextResult());
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void Xlsb_StreamingSnapshot_LateUnsupportedValueLeavesDestinationUntouched(bool sharedStrings) {
            var rows = new SnapshotXlsbRows();
            rows.Values[2][3] = Guid.NewGuid();
            using ExcelDocument document = ExcelDocument.Create();
            document.AddWorksheet("Data");
            var source = new ExcelDirectTabularSource("Data", rows, includeHeaders: true, preserveMissingValues: true);
            using var destination = new MemoryStream();
            byte[] sentinel = Enumerable.Range(0, 64).Select(index => (byte)index).ToArray();
            destination.Write(sentinel, 0, sentinel.Length);

            Assert.False(XlsbNewPackageWriter.TryWriteDirectTabular(document, source, destination,
                CancellationToken.None, sharedStrings));
            Assert.Equal(sentinel, destination.ToArray());
            Assert.Equal(sentinel.Length, destination.Position);
            Assert.Equal(rows.RowCount * rows.ColumnCount, rows.ValueReads);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void Xlsb_StreamingSnapshot_WriteFailureDoesNotContaminateNextSave(bool sharedStrings) {
            using ExcelDocument document = ExcelDocument.Create();
            document.AddWorksheet("Data");
            var failedRows = new SnapshotXlsbRows();
            var failedSource = new ExcelDirectTabularSource("Data", failedRows, includeHeaders: true, preserveMissingValues: true);
            using var failing = new SnapshotXlsbStream(() => throw new IOException("Synthetic destination failure."));
            Assert.Throws<IOException>(() => XlsbNewPackageWriter.TryWriteDirectTabular(
                document, failedSource, failing, CancellationToken.None, sharedStrings));

            var rows = new SnapshotXlsbRows();
            rows.Values[0][0] = "Fresh snapshot";
            rows.Values[0][1] = 123.5D;
            using var destination = new SnapshotXlsbStream(() => { });
            var source = new ExcelDirectTabularSource("Data", rows, includeHeaders: true, preserveMissingValues: true);
            Assert.True(XlsbNewPackageWriter.TryWriteDirectTabular(
                document, source, destination, CancellationToken.None, sharedStrings));
            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(destination.ToArray());
            Assert.True(reader.Read());
            Assert.Equal("Fresh snapshot", reader.GetString(0));
            Assert.Equal(123.5D, reader.GetDouble(1));
            Assert.True(reader.Read());
            Assert.Equal("", reader.GetString(0));
            Assert.True(reader.IsDBNull(1));
        }

        private sealed class SnapshotXlsbRows : IExcelSheetTabularRowSource {
            private readonly string[] _headers = { "Text", "Number", "Date", "Boolean" };
            internal object?[][] Values { get; } = new object?[][] {
                new object?[] { "Zażółć 🚀", 42D, new DateTime(2026, 10, 9, 6, 0, 0), true },
                new object?[] { "", DBNull.Value, null, false },
                new object?[] { null, null, null, null }
            };
            internal int ValueReads { get; private set; }
            internal int HeaderReads { get; private set; }
            public int ColumnCount => _headers.Length;
            public int RowCount => Values.Length;
            public string GetColumnName(int index) { HeaderReads++; return _headers[index]; }
            public Type GetColumnType(int index) => typeof(object);
            public object? GetValue(int rowIndex, int columnIndex) { ValueReads++; return Values[rowIndex][columnIndex]; }
            public bool TryGetBufferedRow(int rowIndex, out object?[]? values) { values = null; return false; }
            public bool TryGetFlatValues(out object?[] values, out int columnCount) {
                values = Array.Empty<object?>();
                columnCount = 0;
                return false;
            }
        }

        private sealed class SnapshotXlsbStream : Stream {
            private readonly MemoryStream _inner = new();
            private readonly Action _onFirstWrite;
            private bool _hasWritten;
            internal SnapshotXlsbStream(Action onFirstWrite) => _onFirstWrite = onFirstWrite;
            public override bool CanRead => false;
            public override bool CanSeek => false;
            public override bool CanWrite => true;
            public override long Length => throw new NotSupportedException();
            public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
            internal byte[] ToArray() => _inner.ToArray();
            public override void Flush() => _inner.Flush();
            public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
            public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
            public override void SetLength(long value) => throw new NotSupportedException();
            public override void Write(byte[] buffer, int offset, int count) {
                if (!_hasWritten) {
                    _hasWritten = true;
                    _onFirstWrite();
                }
                _inner.Write(buffer, offset, count);
            }
            protected override void Dispose(bool disposing) {
                if (disposing) _inner.Dispose();
                base.Dispose(disposing);
            }
        }
    }
}
