using OfficeIMO.Reader;
using OfficeIMO.Reader.Rtf;
using System.Text;
using System.Threading;
using Xunit;

namespace OfficeIMO.Tests;

[Collection("ReaderRegistryNonParallel")]
public sealed class ReaderRtfCancellationTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void Registered_Rtf_Reader_Stops_When_Cancellation_Arrives_During_Input_Read(bool richDocument, bool computeHashes) {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddRtfHandler().Build();
        byte[] input = Encoding.ASCII.GetBytes(@"{\rtf1\ansi " + new string('A', 32_000) + @"\par}");
        using var cancellation = new CancellationTokenSource();
        using var stream = new CancelOnReadStream(input, cancellation);
        var options = new ReaderOptions { ComputeHashes = computeHashes };

        Assert.ThrowsAny<OperationCanceledException>(() => {
            if (richDocument) reader.ReadDocument(stream, "cancel.rtf", options, cancellation.Token);
            else reader.Read(stream, "cancel.rtf", options, cancellation.Token).ToArray();
        });
        Assert.True(stream.BytesRead < input.Length);
        Assert.True(stream.CanRead);
        if (computeHashes) Assert.Equal(0, stream.Position);
    }

    [Fact]
    public void Shared_Stream_Hashing_Hashes_Remaining_Bytes_And_Leaves_Stream_Open() {
        byte[] bytes = Encoding.UTF8.GetBytes("prefix-payload");
        using var stream = new MemoryStream(bytes);
        stream.Position = 7;
        string result = OfficeDocumentAssetHash.ComputeSha256Hex(stream);
        Assert.Equal(OfficeDocumentAssetHash.ComputeSha256Hex(Encoding.UTF8.GetBytes("payload")), result);
        Assert.Equal(bytes.Length, stream.Position);
        Assert.True(stream.CanRead);
    }

    private sealed class CancelOnReadStream : Stream {
        private readonly MemoryStream _source;
        private readonly CancellationTokenSource _cancellation;
        internal CancelOnReadStream(byte[] input, CancellationTokenSource cancellation) {
            _source = new MemoryStream(input, writable: false);
            _cancellation = cancellation;
        }
        internal int BytesRead { get; private set; }
        public override bool CanRead => _source.CanRead;
        public override bool CanSeek => true;
        public override bool CanWrite => false;
        public override long Length => _source.Length;
        public override long Position { get => _source.Position; set => _source.Position = value; }
        public override int Read(byte[] buffer, int offset, int count) {
            int read = _source.Read(buffer, offset, Math.Min(count, 1024));
            BytesRead += read;
            _cancellation.Cancel();
            return read;
        }
        public override long Seek(long offset, SeekOrigin origin) => _source.Seek(offset, origin);
        public override void Flush() { }
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        protected override void Dispose(bool disposing) {
            if (disposing) _source.Dispose();
            base.Dispose(disposing);
        }
    }
}
