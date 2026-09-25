using OfficeIMO.Pdf;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfPageExtractionCancellationTests {
    [Fact]
    public void CancellableInfoSerializationPreservesTextStringBytes() {
        var metadata = new PdfMetadata {
            Title = "Zażółć 😀",
            Author = "OfficeIMO",
            Keywords = "PDF, cancellation"
        };

        Assert.Equal(
            PdfEncoding.Latin1GetBytes(PdfInfoDictionaryBuilder.Build(metadata)),
            PdfInfoDictionaryBuilder.BuildBytesCancellable(metadata, CancellationToken.None));
    }

    [Fact]
    public void LargeEncryptedStreamChunksPreserveCbcAndHonorCancellation() {
        var key = Enumerable.Range(0, 16).Select(static value => (byte)value).ToArray();
        var iv = Enumerable.Range(16, 16).Select(static value => (byte)value).ToArray();
        var plaintext = Enumerable.Range(0, 131072).Select(static value => (byte)value).ToArray();
        byte[] ciphertext = PdfAesCryptography.EncryptNoPadding(key, iv, plaintext, provider: null);

        Assert.Equal(plaintext,
            PdfAesCryptography.DecryptNoPadding(key, iv, ciphertext, provider: null, CancellationToken.None));

        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() =>
            PdfAesCryptography.DecryptNoPadding(key, iv, ciphertext, provider: null, cancellation.Token));
    }

    [Fact]
    public void LargeFalseHeaderScanHonorsCancellation() {
        string source = "%PDF-1.7\n" + string.Concat(Enumerable.Repeat("x obj", 5_000_000));
        byte[] pdf = System.Text.Encoding.ASCII.GetBytes(source);
        var options = new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxObjectParsingTime = TimeSpan.FromSeconds(20) }
        };
        using var cancellation = new CancellationTokenSource();
        using var entered = new ManualResetEventSlim();
        Task parse = Task.Run(() => {
            entered.Set();
            PdfSyntax.ParseObjects(pdf, options, out _, out _, source, cancellation.Token);
        });

        Assert.True(entered.Wait(TimeSpan.FromSeconds(5)));
        cancellation.Cancel();
        Assert.ThrowsAny<OperationCanceledException>(() => parse.GetAwaiter().GetResult());
    }

    [Fact]
    public void LargeHexStringDecodeHonorsCancellation() {
        string hex = new string('A', 40_000_000);
        using var cancellation = new CancellationTokenSource();
        using var entered = new ManualResetEventSlim();
        Task decode = Task.Run(() => {
            entered.Set();
            PdfTextString.DecodeHexBytes(hex, 0, hex.Length, cancellation.Token);
        });

        Assert.True(entered.Wait(TimeSpan.FromSeconds(5)));
        cancellation.Cancel();
        Assert.ThrowsAny<OperationCanceledException>(() => decode.GetAwaiter().GetResult());
    }

    [Fact]
    public void ChunkedPdfTextStringDecodePreservesSplitUtf8Sequences() {
        string expected = new string('a', 4095) + "😀" + new string('b', 4096);
        byte[] bytes = new byte[] { 0xEF, 0xBB, 0xBF }
            .Concat(System.Text.Encoding.UTF8.GetBytes(expected))
            .ToArray();
        using var cancellation = new CancellationTokenSource();

        Assert.Equal(expected, PdfTextString.Decode(bytes, cancellation.Token));
    }

    [Fact]
    public void CancellableUnicodeAndPdfDocStringsKeepTheirDecodedText() {
        string unicode = new string('a', 9000) + "😀";
        byte[] unicodeBytes = PdfTextString.Encode(unicode);
        using var cancellation = new CancellationTokenSource();

        Assert.Equal(unicode, PdfTextString.Decode(unicodeBytes, cancellation.Token));
        Assert.True(PdfJavaScriptStringEncoding.TryDecode(unicodeBytes, out string javaScript, cancellation.Token));
        Assert.Equal(unicode, javaScript);
        Assert.True(PdfDocEncoding.TryDecode(new byte[] { (byte)'A', 0x80 }, out string pdfDoc, cancellation.Token));
        Assert.Equal("A•", pdfDoc);
    }

    [Fact]
    public void JavaScriptDecoderPreservesUnicodeAcrossChunksAndHonorsCancellation() {
        string source = new string('x', 8191) + "😀" + new string('y', 8192);
        byte[] utf8 = new byte[] { 0xEF, 0xBB, 0xBF }
            .Concat(System.Text.Encoding.UTF8.GetBytes(source))
            .ToArray();
        using var validCancellation = new CancellationTokenSource();
        Assert.True(PdfJavaScriptStringEncoding.TryDecode(utf8, out string decoded, validCancellation.Token));
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
    public void PageSelectionStopsWhenTheCallerCancelsEnumeration() {
        byte[] source = CreatePdf();
        using var cancellation = new CancellationTokenSource();

        IEnumerable<int> Selection() {
            yield return 1;
            cancellation.Cancel();
            yield return 1;
        }

        Assert.Throws<OperationCanceledException>(() =>
            PdfPageExtractor.ExtractPages(source, Selection(), options: null, documentFactory: null, cancellation.Token));
    }

    [Fact]
    public void HeaderProbeRejectsAnUnboundedVersionToken() {
        byte[] header = System.Text.Encoding.ASCII.GetBytes("%PDF-" + new string('9', 1_000_000) + "\n");

        Assert.Null(PdfSyntax.GetHeaderVersion(header));
    }

    [Fact]
    public void RawSecurityProbeKeepsEncryptionDetailsAfterLongObjectHeader() {
        byte[] pdf = System.Text.Encoding.ASCII.GetBytes(
            "%PDF-1.7\n" + new string('7', 100_000) + "\n" +
            "7 0 obj" + new string(' ', 2048) + "\n" +
            "<< /Filter /Standard /V 2 /R 3 /Length 128 /P -4 >>\nendobj\n" +
            "trailer\n<< /Encrypt 7 0 R >>\n%%EOF");

        PdfDocumentSecurityInfo security = PdfSyntax.ReadDocumentSecurityInfo(pdf, includeParsedDetails: false);

        Assert.True(security.HasEncryption);
        Assert.Equal("Standard", security.EncryptionFilter);
        Assert.Equal(3, security.EncryptionRevision);
        Assert.Equal(128, security.EncryptionLengthBits);
        Assert.Equal(-4, security.EncryptionPermissions);
    }

    [Fact]
    public void ViewerPreferenceArrayUsesTheConfiguredNestingLimit() {
        string nested = new string('[', 130) + "1" + new string(']', 130);
        byte[] pdf = System.Text.Encoding.ASCII.GetBytes(
            "%PDF-1.7\n" +
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /ViewerPreferences << /Custom " + nested + " >> >>\nendobj\n" +
            "2 0 obj\n<< /Type /Pages /Count 0 /Kids [] >>\nendobj\n" +
            "trailer\n<< /Root 1 0 R >>\n%%EOF");
        var options = new PdfLoadOptions { Limits = new PdfReadLimits { MaxObjectNestingDepth = 160 } };

        PdfReadDocument readback = PdfReadDocument.Open(pdf, options);

        Assert.Equal(nested, readback.ViewerPreferences?.GetValue("Custom"));
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
