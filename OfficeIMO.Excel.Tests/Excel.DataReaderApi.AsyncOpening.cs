using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Theory]
    [InlineData(ExcelFileFormat.Xlsx, true)]
    [InlineData(ExcelFileFormat.Xlsx, false)]
    [InlineData(ExcelFileFormat.Xlsb, true)]
    [InlineData(ExcelFileFormat.Xlsb, false)]
    [InlineData(ExcelFileFormat.Xls, true)]
    [InlineData(ExcelFileFormat.Xls, false)]
    public async Task AsyncOpeningReadsRemainingBytesAndPreservesSourceOwnership(
        ExcelFileFormat format, bool seekable) {
        byte[] workbook = CreateAsyncOpeningWorkbook(format);
        // XLSB also applies this budget to decompressed parts. The prefix alone
        // exceeds the budget, so successful opening proves only remaining bytes are read.
        long inputBudget = workbook.Length * 2L;
        int prefixLength = checked((int)inputBudget + 17);
        byte[] prefixed = new byte[workbook.Length + prefixLength];
        Array.Copy(workbook, 0, prefixed, prefixLength, workbook.Length);
        using var source = new AsyncOpeningReadStream(prefixed, seekable, startPosition: prefixLength);
        using (ExcelWorkbookDataReader reader = await ExcelDocument.OpenDataReaderAsync(
            source,
            new ExcelReadOptions {
                SheetName = "Data",
                MaxInputBytes = inputBudget
            })) {
            Assert.True(source.AsyncReads > 0);
            Assert.Equal("Data", reader.CurrentSheetName);
            Assert.Equal("Name", reader.GetName(0));
            Assert.True(reader.Read());
            Assert.Equal("Alpha", reader.GetString(0));
            Assert.Equal(42, reader.GetInt32(1));
            Assert.False(reader.Read());
        }

        Assert.True(source.CanRead);
        if (seekable) Assert.Equal(prefixLength, source.Position);
    }

    [Theory]
    [InlineData(ExcelFileFormat.Xlsx)]
    [InlineData(ExcelFileFormat.Xlsb)]
    [InlineData(ExcelFileFormat.Xls)]
    public async Task AsyncOpeningPathReturnsDetachedReaderAndClosesSourceFile(ExcelFileFormat format) {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.AsyncOpen.{Guid.NewGuid():N}.{format.ToString().ToLowerInvariant()}");
        try {
            File.WriteAllBytes(path, CreateAsyncOpeningWorkbook(format));
            using ExcelWorkbookDataReader reader = await ExcelDocument.OpenDataReaderAsync(path);
            using (var exclusive = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None)) {
                Assert.True(exclusive.Length > 0);
            }
            Assert.Equal(new[] { "Data", "Next" }, reader.SheetNames);
            Assert.True(reader.Read());
            Assert.Equal("Alpha", reader.GetString(0));
            Assert.True(reader.NextResult());
            Assert.True(reader.Read());
            Assert.Equal("Beta", reader.GetString(0));
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task AsyncOpeningRejectsOverLimitInputWithoutClosingSource(bool seekable) {
        byte[] workbook = CreateAsyncOpeningWorkbook(ExcelFileFormat.Xlsx);
        using var source = new AsyncOpeningReadStream(workbook, seekable);
        await Assert.ThrowsAsync<InvalidDataException>(() => ExcelDocument.OpenDataReaderAsync(
            source, new ExcelReadOptions { MaxInputBytes = workbook.Length - 1 }));
        Assert.True(source.CanRead);
        if (seekable) Assert.Equal(0, source.Position);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task AsyncOpeningObservesBothTokensDuringSourceRead(bool optionsToken) {
        using var cancellation = new CancellationTokenSource();
        using var source = new AsyncOpeningReadStream(CreateAsyncOpeningWorkbook(ExcelFileFormat.Xlsx), seekable: true);
        source.BeforeRead = () => { if (source.AsyncReads == 2) cancellation.Cancel(); };
        var options = new ExcelReadOptions {
            CancellationToken = optionsToken ? cancellation.Token : default
        };
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => ExcelDocument.OpenDataReaderAsync(
            source, options, optionsToken ? default : cancellation.Token));
        Assert.True(source.CanRead);
        Assert.Equal(0, source.Position);
        Assert.Equal(optionsToken ? cancellation.Token : default, options.CancellationToken);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task AsyncOpeningRetainsBothTokensForReaderLifetime(bool optionsToken) {
        using var cancellation = new CancellationTokenSource();
        using var source = new AsyncOpeningReadStream(CreateAsyncOpeningWorkbook(ExcelFileFormat.Xlsx), seekable: true);
        using ExcelWorkbookDataReader reader = await ExcelDocument.OpenDataReaderAsync(
            source,
            new ExcelReadOptions { CancellationToken = optionsToken ? cancellation.Token : default },
            optionsToken ? default : cancellation.Token);
        cancellation.Cancel();
        Assert.ThrowsAny<OperationCanceledException>(() => reader.Read());
        Assert.True(source.CanRead);
    }

    [Fact]
    public async Task AsyncOpeningRestoresSeekableSourceAfterReadOrParseFailure() {
        using var failedRead = new AsyncOpeningReadStream(new byte[256], seekable: true, startPosition: 7);
        failedRead.BeforeRead = () => { if (failedRead.AsyncReads == 2) throw new IOException("Source read failed."); };
        await Assert.ThrowsAsync<IOException>(() => ExcelDocument.OpenDataReaderAsync(failedRead));
        Assert.Equal(7, failedRead.Position);
        Assert.True(failedRead.CanRead);

        using var invalidWorkbook = new AsyncOpeningReadStream(new byte[32], seekable: true, startPosition: 7);
        await Assert.ThrowsAnyAsync<Exception>(() => ExcelDocument.OpenDataReaderAsync(invalidWorkbook));
        Assert.Equal(7, invalidWorkbook.Position);
        Assert.True(invalidWorkbook.CanRead);
    }

    [Fact]
    public async Task AsyncOpeningRejectsCancellationBeforeReadingAndInvalidSources() {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        using var source = new AsyncOpeningReadStream(CreateAsyncOpeningWorkbook(ExcelFileFormat.Xlsx), seekable: true);
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => ExcelDocument.OpenDataReaderAsync(
            source, cancellationToken: cancellation.Token));
        Assert.Equal(0, source.AsyncReads);
        await Assert.ThrowsAsync<ArgumentNullException>(() => ExcelDocument.OpenDataReaderAsync((Stream)null!));
        await Assert.ThrowsAsync<ArgumentException>(() => ExcelDocument.OpenDataReaderAsync(" "));
        await Assert.ThrowsAsync<NotSupportedException>(() => ExcelDocument.OpenDataReaderAsync("input.csv"));
    }

    [Fact]
    public async Task AsyncOpeningAwaitsIncompleteSourceReadAndCopiesOptions() {
        var gate = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        using var source = new AsyncOpeningReadStream(CreateAsyncOpeningWorkbook(ExcelFileFormat.Xlsx), seekable: true) {
            ReadGate = gate.Task
        };
        var options = new ExcelReadOptions { SheetName = "Data" };
        Task<ExcelWorkbookDataReader> opening = ExcelDocument.OpenDataReaderAsync(source, options);
        Assert.False(opening.IsCompleted);
        options.SheetName = "Missing";
        options.MaxInputBytes = 1;
        gate.SetResult(true);
        using ExcelWorkbookDataReader reader = await opening;
        Assert.Equal("Data", reader.CurrentSheetName);
        Assert.True(reader.Read());
        Assert.Equal("Alpha", reader.GetString(0));
        Assert.Equal("Missing", options.SheetName);
        Assert.Equal(1, options.MaxInputBytes);
    }

    private static byte[] CreateAsyncOpeningWorkbook(ExcelFileFormat format) {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Data");
        sheet.CellValue(1, 1, "Name");
        sheet.CellValue(1, 2, "Id");
        sheet.CellValue(2, 1, "Alpha");
        sheet.CellValue(2, 2, 42);
        ExcelSheet next = document.AddWorksheet("Next");
        next.CellValue(1, 1, "Name");
        next.CellValue(2, 1, "Beta");
        return document.ToBytes(format);
    }

    // A caller-owned stream that permits asynchronous input only exposes accidental synchronous I/O.
    private sealed class AsyncOpeningReadStream : Stream {
        private readonly MemoryStream _inner;
        private readonly bool _seekable;
        internal Action? BeforeRead { get; set; }
        internal Task? ReadGate { get; set; }
        internal int AsyncReads { get; private set; }

        internal AsyncOpeningReadStream(byte[] bytes, bool seekable, int startPosition = 0) {
            _inner = new MemoryStream(bytes, writable: false) { Position = startPosition };
            _seekable = seekable;
        }

        public override bool CanRead => _inner.CanRead;
        public override bool CanSeek => _seekable;
        public override bool CanWrite => false;
        public override long Length => _seekable ? _inner.Length : throw new NotSupportedException();
        public override long Position {
            get => _seekable ? _inner.Position : throw new NotSupportedException();
            set { if (!_seekable) throw new NotSupportedException(); _inner.Position = value; }
        }

        public override int Read(byte[] buffer, int offset, int count) =>
            throw new InvalidOperationException("Source requires asynchronous input.");

        public override async Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) {
            AsyncReads++;
            BeforeRead?.Invoke();
            if (ReadGate != null) await ReadGate.ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            return _inner.Read(buffer, offset, Math.Min(count, 113));
        }

        public override void Flush() => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => _seekable ? _inner.Seek(offset, origin) : throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        protected override void Dispose(bool disposing) {
            if (disposing) _inner.Dispose();
            base.Dispose(disposing);
        }
    }
}
