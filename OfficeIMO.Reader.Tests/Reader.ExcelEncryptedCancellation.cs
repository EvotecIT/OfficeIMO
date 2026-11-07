using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Excel;
using OfficeIMO.Reader.Excel;
using OfficeIMO.Tests;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public sealed class ReaderExcelEncryptedCancellationTests {
    [Theory]
    [InlineData("xlsx", false, false)]
    [InlineData("xlsx", true, false)]
    [InlineData("xlsb", false, false)]
    [InlineData("xlsb", true, false)]
    [InlineData("xls", false, false)]
    [InlineData("xls", true, false)]
    [InlineData("xlsx", true, true)]
    public void EncryptedRetryHonorsCancellationBeforeFurtherInput(string format, bool duringRetryRead, bool boundedInput) {
        (byte[] bytes, string password) = EncryptedFixture(format);
        using var cancellation = new CancellationTokenSource();
        using var input = new RetryCancellationStream(bytes, cancellation.Cancel, duringRetryRead);

        OperationCanceledException error = Assert.ThrowsAny<OperationCanceledException>(() => {
            using ExcelDocument document = ExcelReaderAdapter.Load(input, "encrypted." + format,
                new ReaderOptions {
                    OpenPassword = password,
                    Projection = ReaderProjection.Full,
                    MaxInputBytes = boundedInput ? bytes.Length : null
                }, cancellation.Token);
        });

        Assert.Equal(cancellation.Token, error.CancellationToken);
        Assert.True(cancellation.IsCancellationRequested);
        Assert.InRange(input.RetryBytesRead, duringRetryRead ? 1 : 0, duringRetryRead ? 64 : 0);
        Assert.True(input.CanRead);
    }

    [Fact]
    public void EncryptedLoadKeepsConfiguredPackageSizeFailureTerminal() {
        (byte[] bytes, string password) = EncryptedFixture("xlsx");
        using var input = new MemoryStream(bytes, writable: false);
        OfficePackageSecurityException error = Assert.Throws<OfficePackageSecurityException>(() => {
            using ExcelDocument document = ExcelReaderAdapter.Load(input, "encrypted.xlsx",
                new ReaderOptions { OpenPassword = password, MaxInputBytes = bytes.Length - 1 });
        });
        Assert.Equal(OfficePackageSecurityRule.PackageSize, error.Rule);
        Assert.True(input.CanRead);
    }

    [Theory]
    [InlineData("xlsx")]
    [InlineData("xlsb")]
    [InlineData("xls")]
    public async Task EncryptedLoadAsyncHonorsCancellationAfterBuffering(string format) {
        (byte[] bytes, string password) = EncryptedFixture(format);
        using var cancellation = new CancellationTokenSource();
        using var input = new RetryCancellationStream(bytes, cancellation.Cancel, duringRead: true, cancelAtEnd: true);

        OperationCanceledException error = await Assert.ThrowsAnyAsync<OperationCanceledException>(async () => {
            using ExcelDocument document = await ExcelDocument.LoadAsync(input,
                new ExcelLoadOptions { Password = password, AccessMode = DocumentAccessMode.ReadOnly }, cancellation.Token);
        });

        Assert.Equal(cancellation.Token, error.CancellationToken);
        Assert.True(cancellation.IsCancellationRequested);
        Assert.True(input.CanRead);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void EncryptedXlsbRetryKeepsTypedContentAndCallerOwnership(bool pathInput) {
        string path = FixturePath("encrypted-native-control.xlsb");
        using var input = new MemoryStream(File.ReadAllBytes(path), writable: false);
        var options = new ReaderOptions { OpenPassword = "test-password", Projection = ReaderProjection.Full };
        using (ExcelDocument document = pathInput
                   ? ExcelReaderAdapter.Load(path, options, CancellationToken.None)
                   : ExcelReaderAdapter.Load(input, "encrypted.xlsb", options, CancellationToken.None)) {
            Assert.Equal(ExcelFileFormat.Xlsb, document.SourceFormat);
            Assert.Equal("Native encrypted control", document.Sheets[0].Cell(1, 1).GetValue<string>());
            Assert.Equal(12.5d, document.Sheets[0].Cell(2, 1).GetValue<double>());
            Assert.True(document.Sheets[0].Cell(3, 1).GetValue<bool>());
            Assert.Equal("Łódź 00127", document.Sheets[0].Cell(4, 1).GetValue<string>());
            Assert.Equal("A2*2", document.Sheets[0].GetFormulaText(5, 1));
            Assert.Equal(25d, document.Sheets[0].Cell(5, 1).GetValue<double>());
        }
        if (!pathInput) Assert.True(input.CanRead);
    }

    private static (byte[] Bytes, string Password) EncryptedFixture(string format) => format == "xlsx"
        ? (File.ReadAllBytes(FixturePath("agile-aes256-sha512.xlsx")), "hunter2")
        : (File.ReadAllBytes(FixturePath("encrypted-native-control." + format)), "test-password");

    private static string FixturePath(string name) => Path.Combine(RepositoryTestPaths.Find(),
        "OfficeIMO.TestAssets", "Documents", "ExcelEncryptionCorpus", name);

    private sealed class RetryCancellationStream : Stream {
        private readonly MemoryStream _inner;
        private readonly Action _cancel;
        private readonly bool _duringRead;
        private readonly bool _cancelAtEnd;
        private int _passes;
        private bool _cancelled;

        internal RetryCancellationStream(byte[] bytes, Action cancel, bool duringRead, bool cancelAtEnd = false) {
            _inner = new MemoryStream(bytes, writable: false);
            _cancel = cancel;
            _duringRead = duringRead;
            _cancelAtEnd = cancelAtEnd;
        }

        internal long RetryBytesRead { get; private set; }
        public override bool CanRead => _inner.CanRead;
        public override bool CanSeek => true;
        public override bool CanWrite => false;
        public override long Length => _inner.Length;
        public override long Position { get => _inner.Position; set => Seek(value, SeekOrigin.Begin); }

        public override int Read(byte[] buffer, int offset, int count) {
            if (_inner.Position == 0) {
                _passes++;
                if (!_cancelAtEnd && _passes == 2 && _duringRead) Cancel();
            }
            int read = _inner.Read(buffer, offset, Math.Min(count, 64));
            if (_passes >= 2) RetryBytesRead += read;
            if (_cancelAtEnd && read == 0) Cancel();
            return read;
        }

        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) =>
            Task.FromResult(Read(buffer, offset, count));

        public override long Seek(long offset, SeekOrigin origin) {
            long previous = _inner.Position;
            long position = _inner.Seek(offset, origin);
            if (!_duringRead && _passes == 1 && previous != 0 && position == 0) Cancel();
            return position;
        }

        private void Cancel() {
            if (_cancelled) return;
            _cancelled = true;
            _cancel();
        }

        public override void Flush() => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        protected override void Dispose(bool disposing) {
            if (disposing) _inner.Dispose();
            base.Dispose(disposing);
        }
    }
}
