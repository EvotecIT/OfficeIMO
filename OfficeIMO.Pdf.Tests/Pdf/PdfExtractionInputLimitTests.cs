using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfExtractionInputLimitTests {
    [Fact]
    public void ExtractionStreamsRejectOversizedRemainingInputBeforeReading() {
        Action<Stream>[] routes = {
            input => PdfTextExtractor.ExtractAllText(input),
            input => PdfTextExtractor.ExtractTextByPage(input),
            input => PdfTextExtractor.ExtractMarkdown(input),
            input => PdfImageExtractor.ExtractImages(input),
            input => PdfImageExtractor.ExtractImagePlacements(input),
            input => PdfImageExtractor.ExtractImagesByPageRanges(input, PdfPageRange.From(1, 1)),
            input => PdfImageExtractor.ExtractImagePlacementsByPageRanges(input, PdfPageRange.From(1, 1)),
            input => PdfAttachmentExtractor.ExtractAttachments(input)
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
    public void ExtractionPathsRejectOversizedFilesBeforeBuffering() {
        string path = Path.Combine(Path.GetTempPath(), "officeimo-pdf-extraction-limit-" + Guid.NewGuid().ToString("N") + ".pdf");
        string root = Path.Combine(Path.GetTempPath(), "officeimo-pdf-extraction-output-" + Guid.NewGuid().ToString("N"));
        try {
            using (var file = new FileStream(path, FileMode.CreateNew, FileAccess.Write)) {
                file.SetLength(PdfLoadOptions.Default.Limits.MaxInputBytes + 1);
            }

            Action<string>[] routes = {
                input => PdfTextExtractor.ExtractAllText(input),
                input => PdfImageExtractor.ExtractImages(input),
                input => PdfAttachmentExtractor.ExtractAttachments(input)
            };
            foreach (Action<string> route in routes) {
                Assert.Equal(PdfReadLimitKind.InputBytes,
                    Assert.Throws<PdfReadLimitException>(() => route(path)).Kind);
            }

            string textDirectory = Path.Combine(root, "text");
            string imageDirectory = Path.Combine(root, "images");
            string attachmentDirectory = Path.Combine(root, "attachments");
            Assert.Throws<PdfReadLimitException>(() => PdfTextExtractor.ExtractTextByPage(path, textDirectory));
            Assert.Throws<PdfReadLimitException>(() => PdfImageExtractor.ExtractImages(path, imageDirectory));
            Assert.Throws<PdfReadLimitException>(() => PdfAttachmentExtractor.ExtractAttachments(path, attachmentDirectory));
            Assert.False(Directory.Exists(textDirectory));
            Assert.False(Directory.Exists(imageDirectory));
            Assert.False(Directory.Exists(attachmentDirectory));
        } finally {
            File.Delete(path);
            if (Directory.Exists(root)) Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void OversizedTextInputLeavesOutputStreamUntouched() {
        using var input = new OversizedLengthStream(PdfLoadOptions.Default.Limits.MaxInputBytes + 1);
        using var output = new MemoryStream();

        Assert.Equal(PdfReadLimitKind.InputBytes,
            Assert.Throws<PdfReadLimitException>(() => PdfTextExtractor.ExtractAllText(input, output)).Kind);
        Assert.Equal(0, output.Length);
        Assert.False(input.WasRead);
    }

    [Fact]
    public void TextExtractionUsesCurrentStreamPositionAcrossReadModels() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Current position marker"))
            .ToBytes();
        byte[] prefixed = new byte[pdf.Length + 5];
        Buffer.BlockCopy(pdf, 0, prefixed, 5, pdf.Length);

        Action<Stream>[] routes = {
            input => Assert.Contains("Current position marker", PdfTextExtractor.ExtractAllText(input), StringComparison.Ordinal),
            input => Assert.Single(PdfTextExtractor.ExtractTextByPage(input)),
            input => Assert.NotEmpty(PdfTextExtractor.ExtractMarkdown(input)),
            input => Assert.Single(PdfTextExtractor.ExtractStructuredByPage(input))
        };
        foreach (Action<Stream> route in routes) {
            using var input = new MemoryStream(prefixed);
            input.Position = 5;
            route(input);
            Assert.Equal(input.Length, input.Position);
        }
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
