using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfPageExtractionInputLimitTests {
    [Fact]
    public void PageExtractionStreamsRejectOversizedRemainingInputBeforeReading() {
        Action<Stream>[] routes = {
            input => PdfPageExtractor.ExtractPages(input, 1),
            input => PdfPageExtractor.ExtractPageRange(input, 1, 1),
            input => PdfPageExtractor.ExtractPageRanges(input, PdfPageRange.From(1, 1)),
            input => PdfPageExtractor.SplitPages(input),
            input => PdfPageExtractor.SplitPageRanges(input, PdfPageRange.From(1, 1))
        };

        foreach (Action<Stream> route in routes) {
            using var input = new OversizedLengthStream(PdfLoadOptions.Default.Limits.MaxInputBytes + 4);
            input.Position = 3;

            PdfReadLimitException error = Assert.Throws<PdfReadLimitException>(() => route(input));
            Assert.Equal(PdfReadLimitKind.InputBytes, error.Kind);
            Assert.Equal(3, input.Position);
            Assert.False(input.WasRead);
        }
    }

    [Fact]
    public void PageExtractionPathsRejectOversizedInputWithoutCreatingOutputs() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-page-extraction-limit-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(root, "input.pdf");
        string outputPath = Path.Combine(root, "output.pdf");
        string outputDirectory = Path.Combine(root, "split");
        try {
            Directory.CreateDirectory(root);
            using (var file = new FileStream(inputPath, FileMode.CreateNew, FileAccess.Write)) {
                file.SetLength(PdfLoadOptions.Default.Limits.MaxInputBytes + 1);
            }

            Action<string>[] routes = {
                input => PdfPageExtractor.ExtractPages(input, 1),
                input => PdfPageExtractor.ExtractPageRange(input, 1, 1),
                input => PdfPageExtractor.ExtractPageRanges(input, PdfPageRange.From(1, 1)),
                input => PdfPageExtractor.SplitPages(input),
                input => PdfPageExtractor.SplitPageRanges(input, PdfPageRange.From(1, 1))
            };
            foreach (Action<string> route in routes) {
                Assert.Equal(PdfReadLimitKind.InputBytes,
                    Assert.Throws<PdfReadLimitException>(() => route(inputPath)).Kind);
            }

            Assert.Equal(PdfReadLimitKind.InputBytes,
                Assert.Throws<PdfReadLimitException>(() => PdfPageExtractor.ExtractPages(inputPath, outputPath, 1)).Kind);
            Assert.Equal(PdfReadLimitKind.InputBytes,
                Assert.Throws<PdfReadLimitException>(() => PdfPageExtractor.SplitPages(inputPath, outputDirectory)).Kind);
            Assert.False(File.Exists(outputPath));
            Assert.False(Directory.Exists(outputDirectory));
        } finally {
            if (Directory.Exists(root)) Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void PageExtractionReadsFromCurrentStreamPosition() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Page extraction current position"))
            .ToBytes();
        byte[] prefixed = new byte[pdf.Length + 5];
        Buffer.BlockCopy(pdf, 0, prefixed, 5, pdf.Length);
        using var input = new MemoryStream(prefixed);
        input.Position = 5;

        byte[] extracted = PdfPageExtractor.ExtractPages(input, 1);

        Assert.Contains("Page extraction current position", PdfTextExtractor.ExtractAllText(extracted), StringComparison.Ordinal);
        Assert.Equal(input.Length, input.Position);
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
}
