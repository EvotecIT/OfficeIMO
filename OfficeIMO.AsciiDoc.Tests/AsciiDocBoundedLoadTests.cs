using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.AsciiDoc;
using Xunit;

namespace OfficeIMO.AsciiDoc.Tests;

public sealed class AsciiDocBoundedLoadTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task SuppliedEncodingControlsAllLoadPaths(bool matchingBom) {
        Encoding encoding = matchingBom ? new UTF8Encoding(true) : Encoding.GetEncoding(28591);
        byte[] content = encoding.GetBytes("Café");
        byte[] bytes = new byte[] { 0xef, 0xbb, 0xbf }.Concat(content).ToArray();
        string expected = matchingBom ? "Café" : encoding.GetString(bytes);
        string path = Path.Combine(Path.GetTempPath(), "officeimo-explicit-encoding-" + Guid.NewGuid().ToString("N") + ".adoc");
        File.WriteAllBytes(path, bytes);
        try {
            using var stream = new MemoryStream(bytes);
            stream.Position = 2;
            Assert.Equal(expected, AsciiDocDocument.Load(stream, encoding: encoding).ToAsciiDoc());
            Assert.Equal(2, stream.Position);
            Assert.Equal(expected, (await AsciiDocDocument.LoadAsync(stream, encoding: encoding)).ToAsciiDoc());
            Assert.Equal(2, stream.Position);
            Assert.True(stream.CanRead);
            Assert.Equal(expected, AsciiDocDocument.Load(path, encoding: encoding).ToAsciiDoc());
            Assert.Equal(expected, (await AsciiDocDocument.LoadAsync(path, encoding: encoding)).ToAsciiDoc());
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("utf8")]
    [InlineData("utf16le")]
    [InlineData("utf16be")]
    [InlineData("utf32le")]
    [InlineData("utf32be")]
    public async Task BomDecodingAgreesAcrossPathStreamAndAsyncLoads(string encodingName) {
        Encoding encoding = encodingName switch {
            "utf8" => new UTF8Encoding(true),
            "utf16le" => new UnicodeEncoding(false, true),
            "utf16be" => new UnicodeEncoding(true, true),
            "utf32le" => new UTF32Encoding(false, true),
            _ => new UTF32Encoding(true, true)
        };
        const string source = "= Title\r\n\r\nCafé 😀\r\n";
        byte[] bytes = encoding.GetPreamble().Concat(encoding.GetBytes(source)).ToArray();
        string path = Path.Combine(Path.GetTempPath(), "officeimo-bom-" + Guid.NewGuid().ToString("N") + ".adoc");
        File.WriteAllBytes(path, bytes);
        try {
            using var stream = new MemoryStream(bytes);
            stream.Position = 3;
            Assert.Equal(source, AsciiDocDocument.Load(stream).ToAsciiDoc());
            Assert.Equal(3, stream.Position);
            Assert.Equal(source, (await AsciiDocDocument.LoadAsync(stream)).ToAsciiDoc());
            Assert.Equal(3, stream.Position);
            Assert.Equal("Title", AsciiDocDocument.Load(path).BlocksOfType<AsciiDocHeading>().Single().Title);
            Assert.Equal(source, (await AsciiDocDocument.LoadAsync(path)).ToAsciiDoc());
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task OversizedNonSeekableInputStopsAfterBoundedRead(bool asynchronous) {
        using var stream = new CountingTextStream(1024 * 1024);
        var options = new AsciiDocParseOptions { MaximumInputLength = 10 };
        if (asynchronous) {
            await Assert.ThrowsAsync<InvalidDataException>(() => AsciiDocDocument.LoadAsync(stream, options));
        } else {
            Assert.Throws<InvalidDataException>(() => AsciiDocDocument.Load(stream, options));
        }
        Assert.InRange(stream.BytesRead, 11, 8192);
        Assert.True(stream.CanRead);
    }

    private sealed class CountingTextStream : Stream {
        private readonly int _length;
        internal CountingTextStream(int length) => _length = length;
        internal int BytesRead { get; private set; }
        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override int Read(byte[] buffer, int offset, int count) {
            int read = Math.Min(count, _length - BytesRead);
            for (int index = 0; index < read; index++) buffer[offset + index] = (byte)'a';
            BytesRead += read;
            return read;
        }
        public override void Flush() => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }

    [Fact]
    public async Task AsyncLoadForwardsCancellationToTheCallerStream() {
        using var cancellation = new CancellationTokenSource();
        using var stream = new CancellationProbeStream(cancellation);
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => AsciiDocDocument.LoadAsync(stream, cancellationToken: cancellation.Token));
        Assert.Equal(cancellation.Token, stream.ObservedToken);
        Assert.True(stream.CanRead);
    }

    private sealed class CancellationProbeStream : Stream {
        private readonly CancellationTokenSource _cancellation;
        internal CancellationProbeStream(CancellationTokenSource cancellation) => _cancellation = cancellation;
        internal CancellationToken ObservedToken { get; private set; }
        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) {
            ObservedToken = cancellationToken;
            _cancellation.Cancel();
            cancellationToken.ThrowIfCancellationRequested();
            return Task.FromResult(0);
        }
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override void Flush() => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }
}
