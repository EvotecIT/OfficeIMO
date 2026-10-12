using OfficeIMO.Excel;
using OfficeIMO.Excel.Xlsb.Write;
using System.Threading;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData(false, false)]
        [InlineData(true, false)]
        [InlineData(false, true)]
        [InlineData(true, true)]
        public void Xlsb_StreamingSnapshot_ReadsEachFieldOnceBeforeDestinationMutation(bool sharedStrings, bool aboveCaptureLimit) {
            var rows = new SnapshotXlsbRows(aboveCaptureLimit);
            DateTime expectedDate = (DateTime)rows.Values[0][2]!;
            using ExcelDocument document = ExcelDocument.Create();
            document.DateSystem = aboveCaptureLimit ? ExcelDateSystem.NineteenFour : ExcelDateSystem.NineteenHundred;
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
            destination.Flush();
            byte[] package = destination.ToArray();
            AssertXlsbStringStorage(package, sharedStrings, rows.ColumnCount + 2,
                Enumerable.Range(0, rows.ColumnCount).Select(rows.ColumnNameAt).Concat(new[] { "Zażółć 🚀", "" }).ToArray());
            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(package);
            Assert.Equal(rows.ColumnCount, reader.FieldCount);
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
            for (int row = 3; row < rows.RowCount; row++) {
                Assert.True(reader.Read());
                Assert.True(reader.IsDBNull(0));
                Assert.True(reader.IsDBNull(rows.ColumnCount - 1));
            }
            Assert.False(reader.Read());
            Assert.False(reader.NextResult());
        }

        [Theory]
        [InlineData(false, false)]
        [InlineData(true, false)]
        [InlineData(false, true)]
        [InlineData(true, true)]
        public void Xlsb_StreamingSnapshot_LateUnsupportedValueLeavesDestinationUntouched(bool sharedStrings, bool aboveCaptureLimit) {
            var rows = new SnapshotXlsbRows(aboveCaptureLimit) { LastValue = Guid.NewGuid() };
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
        [InlineData(false, false)]
        [InlineData(true, false)]
        [InlineData(false, true)]
        [InlineData(true, true)]
        public void Xlsb_StreamingSnapshot_WriteFailureDoesNotContaminateNextSave(bool sharedStrings, bool aboveCaptureLimit) {
            using ExcelDocument document = ExcelDocument.Create();
            document.AddWorksheet("Data");
            var failedRows = new SnapshotXlsbRows(aboveCaptureLimit);
            var failedSource = new ExcelDirectTabularSource("Data", failedRows, includeHeaders: true, preserveMissingValues: true);
            using var failing = new SnapshotXlsbStream(() => throw new IOException("Synthetic destination failure."));
            Assert.Throws<IOException>(() => XlsbNewPackageWriter.TryWriteDirectTabular(
                document, failedSource, failing, CancellationToken.None, sharedStrings));

            var rows = new SnapshotXlsbRows(aboveCaptureLimit);
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

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void Xlsb_StagedPackage_LateOverlongTextLeavesDestinationUntouched(bool sharedStrings) {
            var rows = new SnapshotXlsbRows(aboveCaptureLimit: true) { LastValue = new string('X', 32_768) };
            using ExcelDocument document = ExcelDocument.Create();
            document.AddWorksheet("Data");
            var source = new ExcelDirectTabularSource("Data", rows, includeHeaders: true, preserveMissingValues: true);
            using var destination = new MemoryStream();
            byte[] sentinel = Enumerable.Range(0, 64).Select(index => (byte)index).ToArray();
            destination.Write(sentinel, 0, sentinel.Length);

            Assert.Throws<ArgumentException>(() => XlsbNewPackageWriter.TryWriteDirectTabular(
                document, source, destination, CancellationToken.None, sharedStrings));

            Assert.Equal(sentinel, destination.ToArray());
            Assert.Equal(sentinel.Length, destination.Position);
            Assert.Equal(rows.RowCount * rows.ColumnCount, rows.ValueReads);
        }

        [Theory]
        [InlineData(false, false)]
        [InlineData(true, false)]
        [InlineData(false, true)]
        [InlineData(true, true)]
        public void Xlsb_StagedPackage_CancellationDuringInputLeavesDestinationUntouched(bool sharedStrings, bool cancelAtLastValue) {
            using var cancellation = new CancellationTokenSource();
            var rows = new SnapshotXlsbRows(aboveCaptureLimit: true);
            int cancelAfterReads = cancelAtLastValue
                ? rows.RowCount * rows.ColumnCount
                : 1024 * rows.ColumnCount;
            rows.OnValueRead = reads => { if (reads == cancelAfterReads) cancellation.Cancel(); };
            using ExcelDocument document = ExcelDocument.Create();
            document.AddWorksheet("Data");
            var source = new ExcelDirectTabularSource("Data", rows, includeHeaders: true, preserveMissingValues: true);
            using var destination = new MemoryStream();
            byte[] sentinel = Enumerable.Range(0, 64).Select(index => (byte)index).ToArray();
            destination.Write(sentinel, 0, sentinel.Length);

            OperationCanceledException exception = Assert.Throws<OperationCanceledException>(() =>
                XlsbNewPackageWriter.TryWriteDirectTabular(document, source, destination, cancellation.Token, sharedStrings));

            Assert.Equal(cancellation.Token, exception.CancellationToken);
            Assert.Equal(sentinel, destination.ToArray());
            Assert.Equal(sentinel.Length, destination.Position);
            if (cancelAtLastValue) {
                Assert.Equal(cancelAfterReads, rows.ValueReads);
            } else {
                Assert.InRange(rows.ValueReads, cancelAfterReads, rows.RowCount * rows.ColumnCount - 1);
            }
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void Xlsb_StagedPackage_WithoutHeadersPreservesTextAndMissingValues(bool sharedStrings) {
            var rows = new SnapshotXlsbRows(aboveCaptureLimit: true);
            DateTime expectedDate = (DateTime)rows.Values[0][2]!;
            using ExcelDocument document = ExcelDocument.Create();
            document.AddWorksheet("Data");
            var source = new ExcelDirectTabularSource("Data", rows, includeHeaders: false, preserveMissingValues: true);
            using var destination = new MemoryStream();
            byte[] existingContent = new byte[256 * 1024];
            destination.Write(existingContent, 0, existingContent.Length);

            Assert.True(XlsbNewPackageWriter.TryWriteDirectTabular(
                document, source, destination, CancellationToken.None, sharedStrings));

            Assert.Equal(rows.RowCount * rows.ColumnCount, rows.ValueReads);
            Assert.Equal(0, rows.HeaderReads);
            Assert.Equal(destination.Length, destination.Position);
            byte[] package = destination.ToArray();
            AssertXlsbStringStorage(package, sharedStrings, 2, new[] { "Zażółć 🚀", "" });
            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(package, new ExcelReadOptions { HasHeaderRow = false });
            Assert.Equal(rows.ColumnCount, reader.FieldCount);
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
            for (int row = 2; row < rows.RowCount; row++) {
                Assert.True(reader.Read());
                Assert.True(reader.IsDBNull(0));
                Assert.True(reader.IsDBNull(rows.ColumnCount - 1));
            }
            Assert.False(reader.Read());
            Assert.False(reader.NextResult());
        }

        private sealed class SnapshotXlsbRows : IExcelSheetTabularRowSource {
            private readonly string[] _headers = { "Text", "Number", "Date", "Boolean" };
            private readonly bool _aboveCaptureLimit;
            internal SnapshotXlsbRows(bool aboveCaptureLimit = false) => _aboveCaptureLimit = aboveCaptureLimit;
            internal object?[][] Values { get; } = new object?[][] {
                new object?[] { "Zażółć 🚀", 42D, new DateTime(2026, 10, 9, 6, 0, 0), true },
                new object?[] { "", DBNull.Value, null, false },
                new object?[] { null, null, null, null }
            };
            internal object? LastValue { get; set; }
            internal Action<int>? OnValueRead { get; set; }
            internal int ValueReads { get; private set; }
            internal int HeaderReads { get; private set; }
            // Cross the capture limit without materializing a million values or a large worksheet payload.
            public int ColumnCount => _aboveCaptureLimit ? 512 : _headers.Length;
            public int RowCount => _aboveCaptureLimit ? 2049 : Values.Length;
            internal string ColumnNameAt(int index) => index < _headers.Length ? _headers[index] : "Column " + (index + 1);
            public string GetColumnName(int index) { HeaderReads++; return ColumnNameAt(index); }
            public Type GetColumnType(int index) => typeof(object);
            public object? GetValue(int rowIndex, int columnIndex) {
                ValueReads++;
                OnValueRead?.Invoke(ValueReads);
                if (rowIndex == RowCount - 1 && columnIndex == ColumnCount - 1) return LastValue;
                return rowIndex < Values.Length && columnIndex < _headers.Length ? Values[rowIndex][columnIndex] : null;
            }
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
