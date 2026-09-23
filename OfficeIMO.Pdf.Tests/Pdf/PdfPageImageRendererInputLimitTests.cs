using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfPageImageRendererInputLimitTests {
    [Fact]
    public void PageRenderStreamEntrypointsRejectOversizedRemainingInputBeforeReading() {
        Action<Stream>[] routes = {
            input => PdfPageImageRenderer.RenderPage(input),
            input => PdfPageImageRenderer.RenderPageAsPng(input),
            input => PdfPageImageRenderer.RenderPageAsSvg(input)
        };

        foreach (Action<Stream> route in routes) {
            using var input = CreateOversizedStream();

            PdfReadLimitException error = Assert.Throws<PdfReadLimitException>(() => route(input));

            Assert.Equal(PdfReadLimitKind.InputBytes, error.Kind);
            Assert.Equal(3, input.Position);
            Assert.False(input.WasRead);
        }
    }

    [Fact]
    public void PageRenderSourcesHonorCustomLimitsForNonSeekableStreamsAndPaths() {
        var options = new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxInputBytes = 1024 }
        };
        byte[] bytes = new byte[4096];

        using (var input = new ChunkedNonSeekableStream(bytes, 256)) {
            PdfReadLimitException error = Assert.Throws<PdfReadLimitException>(() =>
                PdfPageImageRenderer.RenderPage(input, 1, options));
            Assert.Equal(PdfReadLimitKind.InputBytes, error.Kind);
            Assert.InRange(input.BytesRead, 1025, 1280);
        }

        string path = Path.Combine(Path.GetTempPath(), "officeimo-page-render-limit-" + Guid.NewGuid().ToString("N") + ".pdf");
        try {
            File.WriteAllBytes(path, bytes);
            PdfReadLimitException error = Assert.Throws<PdfReadLimitException>(() =>
                PdfPageImageRenderer.RenderPage(path, 1, options));
            Assert.Equal(PdfReadLimitKind.InputBytes, error.Kind);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void PageRenderStreamEntrypointsContinueFromCallerPosition() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("render source"))
            .ToBytes();

        using (var input = CreatePrefixedStream(source)) {
            OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(input);
            Assert.NotNull(drawing);
            Assert.Equal(input.Length, input.Position);
        }

        using (var input = CreatePrefixedStream(source)) {
            byte[] png = PdfPageImageRenderer.RenderPageAsPng(input);
            Assert.Equal(new byte[] { 137, 80, 78, 71 }, png.Take(4).ToArray());
            Assert.Equal(input.Length, input.Position);
        }

        using (var input = CreatePrefixedStream(source)) {
            byte[] svg = PdfPageImageRenderer.RenderPageAsSvg(input);
            Assert.Contains("<svg", System.Text.Encoding.UTF8.GetString(svg), StringComparison.Ordinal);
            Assert.Equal(input.Length, input.Position);
        }
    }

    private static MemoryStream CreatePrefixedStream(byte[] pdf) {
        var stream = new MemoryStream();
        stream.Write(new byte[] { 9, 8, 7 }, 0, 3);
        stream.Write(pdf, 0, pdf.Length);
        stream.Position = 3;
        return stream;
    }

    private static OversizedLengthStream CreateOversizedStream() {
        var stream = new OversizedLengthStream(PdfLoadOptions.Default.Limits.MaxInputBytes + 4) { Position = 3 };
        return stream;
    }

    private sealed class OversizedLengthStream : Stream {
        private readonly long _length;
        private long _position;

        internal OversizedLengthStream(long length) => _length = length;
        internal bool WasRead { get; private set; }
        public override bool CanRead => true;
        public override bool CanSeek => true;
        public override bool CanWrite => false;
        public override long Length => _length;
        public override long Position { get => _position; set => _position = value; }
        public override int Read(byte[] buffer, int offset, int count) {
            WasRead = true;
            throw new InvalidOperationException("Oversized input should be rejected before reading.");
        }
        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }

    private sealed class ChunkedNonSeekableStream : Stream {
        private readonly byte[] _bytes;
        private readonly int _chunkSize;
        private int _position;

        internal ChunkedNonSeekableStream(byte[] bytes, int chunkSize) {
            _bytes = bytes;
            _chunkSize = chunkSize;
        }

        internal int BytesRead => _position;
        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => _position; set => throw new NotSupportedException(); }
        public override int Read(byte[] buffer, int offset, int count) {
            int remaining = _bytes.Length - _position;
            if (remaining <= 0) return 0;
            int read = Math.Min(Math.Min(count, _chunkSize), remaining);
            Array.Copy(_bytes, _position, buffer, offset, read);
            _position += read;
            return read;
        }
        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }
}
