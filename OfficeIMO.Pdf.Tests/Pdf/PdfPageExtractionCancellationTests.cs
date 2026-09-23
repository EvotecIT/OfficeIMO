using OfficeIMO.Pdf;
using System.Threading;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfPageExtractionCancellationTests {
    [Fact]
    public void JavaScriptDecoderPreservesUnicodeAcrossChunksAndHonorsCancellation() {
        string source = new string('x', 8191) + "😀" + new string('y', 8192);
        byte[] utf8 = new byte[] { 0xEF, 0xBB, 0xBF }
            .Concat(System.Text.Encoding.UTF8.GetBytes(source))
            .ToArray();
        Assert.True(PdfJavaScriptStringEncoding.TryDecode(utf8, out string decoded));
        Assert.Equal(source, decoded);

        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() =>
            PdfJavaScriptStringEncoding.TryDecode(utf8, out _, cancellation.Token));
    }

    [Fact]
    public void ExtractPagesReadsNonSeekableInputFromCurrentPositionWithExactLimit() {
        byte[] pdf = CreatePdf();
        byte[] inputBytes = new byte[pdf.Length + 3];
        Buffer.BlockCopy(pdf, 0, inputBytes, 3, pdf.Length);
        var options = new PdfLoadOptions { Limits = new PdfReadLimits { MaxInputBytes = pdf.Length } };
        using var input = new ChunkedNonSeekableStream(inputBytes, 3, maximumChunkSize: 7);

        byte[] extracted = PdfPageExtractor.ExtractPages(input, new[] { 1 }, options, CancellationToken.None);

        Assert.Contains("Page extraction cancellation", PdfTextExtractor.ExtractAllText(extracted), StringComparison.Ordinal);
        Assert.Equal(inputBytes.Length, input.BytesRead);
    }

    [Fact]
    public void ExtractPagesRejectsOversizedNonSeekableInputWithoutChangingOutput() {
        byte[] pdf = CreatePdf();
        var options = new PdfLoadOptions { Limits = new PdfReadLimits { MaxInputBytes = pdf.Length - 1 } };
        using var input = new ChunkedNonSeekableStream(pdf, 0, maximumChunkSize: 7);
        using var output = new MemoryStream();
        output.WriteByte(42);

        PdfReadLimitException error = Assert.Throws<PdfReadLimitException>(() =>
            PdfPageExtractor.ExtractPages(input, output, new[] { 1 }, options, CancellationToken.None));

        Assert.Equal(PdfReadLimitKind.InputBytes, error.Kind);
        Assert.Equal(pdf.Length - 1, error.Limit);
        Assert.Equal(new byte[] { 42 }, output.ToArray());
    }

    [Fact]
    public void ExtractPagesStopsAfterNonSeekableInputCancelsWithoutChangingOutput() {
        byte[] pdf = CreatePdf();
        using var cancellation = new CancellationTokenSource();
        using var input = new ChunkedNonSeekableStream(pdf, 0, maximumChunkSize: 7, cancellation);
        using var output = new MemoryStream();
        output.WriteByte(42);

        Assert.Throws<OperationCanceledException>(() =>
            PdfPageExtractor.ExtractPages(input, output, new[] { 1 }, options: null, cancellation.Token));

        Assert.InRange(input.BytesRead, 1, pdf.Length - 1);
        Assert.Equal(new byte[] { 42 }, output.ToArray());
    }

    [Fact]
    public void DocumentPageExtractionHonorsCancellationAndDoesNotChangeSource() {
        byte[] source = CreatePdf();
        using var cancellation = new CancellationTokenSource();
        PdfDocument document = PdfDocument.Load(source);
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() => document.Pages.Extract(cancellation.Token, 1));
        Assert.Equal(source, document.ToBytes());

        PdfDocument extracted = document.Pages.Extract(CancellationToken.None, 1);
        Assert.Contains("Page extraction cancellation", PdfTextExtractor.ExtractAllText(extracted.ToBytes()), StringComparison.Ordinal);
    }

    [Fact]
    public void ArtifactCaptureKeepsExactDigestAndHonorsCancellation() {
        byte[] pdf = CreatePdf();
        using var cancellation = new CancellationTokenSource();

        PdfArtifactSnapshot expected = PdfArtifactSnapshot.CaptureKnownPageCount(pdf, 1);
        PdfArtifactSnapshot actual = PdfArtifactSnapshot.CaptureKnownPageCount(pdf, 1, cancellation.Token);
        Assert.Equal(expected.Sha256, actual.Sha256);
        Assert.Equal(expected.ByteCount, actual.ByteCount);

        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() =>
            PdfArtifactSnapshot.CaptureKnownPageCount(pdf, 1, cancellation.Token));
    }

    private static byte[] CreatePdf() => PdfDocument.Create()
        .Paragraph(paragraph => paragraph.Text("Page extraction cancellation"))
        .ToBytes();

    private sealed class ChunkedNonSeekableStream : Stream {
        private readonly byte[] _bytes;
        private readonly int _maximumChunkSize;
        private readonly CancellationTokenSource? _cancelAfterFirstRead;
        private int _position;

        internal ChunkedNonSeekableStream(byte[] bytes, int startPosition, int maximumChunkSize,
            CancellationTokenSource? cancelAfterFirstRead = null) {
            _bytes = bytes;
            _position = startPosition;
            _maximumChunkSize = maximumChunkSize;
            _cancelAfterFirstRead = cancelAfterFirstRead;
        }

        internal int BytesRead => _position;
        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override int Read(byte[] buffer, int offset, int count) {
            int available = _bytes.Length - _position;
            if (available <= 0) return 0;
            int read = Math.Min(Math.Min(count, _maximumChunkSize), available);
            Buffer.BlockCopy(_bytes, _position, buffer, offset, read);
            _position += read;
            _cancelAfterFirstRead?.Cancel();
            return read;
        }
        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }
}
