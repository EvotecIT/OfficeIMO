using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelEncryptedLoadCancellationTests {
    [Theory]
    [InlineData(ExcelFileFormat.Xlsx)]
    [InlineData(ExcelFileFormat.Xlsb)]
    [InlineData(ExcelFileFormat.Xls)]
    public async Task EncryptedLoadObservesCancellationAfterInputEnds(ExcelFileFormat format) {
        const string password = "test-password";
        byte[] encrypted;
        using (var document = ExcelDocument.Create()) {
            document.AddSheet("Data").Cell(1, 1).SetValue("Encrypted value");
            encrypted = document.ToBytes(new ExcelSaveOptions { Format = format, Password = password });
        }

        using (var control = new MemoryStream(encrypted, writable: false))
        using (var loaded = await ExcelDocument.LoadAsync(control, new ExcelLoadOptions { Password = password })) {
            Assert.Equal("Encrypted value", loaded.Sheets[0].Cell(1, 1).GetValue<string>());
            Assert.Equal(format, loaded.SourceFormat);
        }

        using var cancellation = new CancellationTokenSource();
        using var source = new CancelAtEndOfInputStream(encrypted, cancellation) { Position = 3 };
        OperationCanceledException error = await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            ExcelDocument.LoadAsync(source, new ExcelLoadOptions { Password = password }, cancellation.Token));
        Assert.Equal(cancellation.Token, error.CancellationToken);
        Assert.True(source.EndObserved);
        Assert.Equal(encrypted.Length, source.BytesRead);
        Assert.True(source.CanRead);
        Assert.Equal(3, source.Position);
    }

    private sealed class CancelAtEndOfInputStream : Stream {
        private readonly MemoryStream _input;
        private readonly CancellationTokenSource _cancellation;

        internal CancelAtEndOfInputStream(byte[] bytes, CancellationTokenSource cancellation) {
            _input = new MemoryStream(bytes, writable: false);
            _cancellation = cancellation;
        }

        internal bool EndObserved { get; private set; }
        internal int BytesRead { get; private set; }
        public override bool CanRead => _input.CanRead;
        public override bool CanSeek => true;
        public override bool CanWrite => false;
        public override long Length => _input.Length;
        public override long Position { get => _input.Position; set => _input.Position = value; }

        public override int Read(byte[] buffer, int offset, int count) {
            int read = _input.Read(buffer, offset, count);
            BytesRead += read;
            if (read == 0 && count > 0 && _input.Position == _input.Length) {
                EndObserved = true;
                _cancellation.Cancel();
            }
            return read;
        }

        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) {
            cancellationToken.ThrowIfCancellationRequested();
            return Task.FromResult(Read(buffer, offset, count));
        }

        public override long Seek(long offset, SeekOrigin origin) => _input.Seek(offset, origin);
        public override void Flush() { }
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        protected override void Dispose(bool disposing) {
            if (disposing) _input.Dispose();
            base.Dispose(disposing);
        }
    }
}
