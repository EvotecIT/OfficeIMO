using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfFormStampInputLimitTests {
    private static readonly IReadOnlyDictionary<string, string> EmptyFields =
        new Dictionary<string, string>();

    [Fact]
    public void FormAndStampStreamsRejectOversizedRemainingInputBeforeReading() {
        Action<Stream>[] routes = {
            input => PdfFormFiller.FillFields(input, EmptyFields),
            input => PdfFormFiller.FlattenFields(input),
            input => PdfFormFiller.FillAndFlattenFields(input, EmptyFields),
            input => PdfStamper.StampText(input, "stamp"),
            input => PdfStamper.WatermarkText(input, "watermark"),
            input => PdfStamper.StampImage(input, Array.Empty<byte>()),
            input => PdfStamper.WatermarkImage(input, Array.Empty<byte>())
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
    public void FormAndStampPathsRejectOversizedInputWithoutWritingOutputs() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-form-stamp-limit-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(root, "input.pdf");
        string formOutput = Path.Combine(root, "form.pdf");
        string stampOutput = Path.Combine(root, "stamp.pdf");
        try {
            Directory.CreateDirectory(root);
            using (var file = new FileStream(inputPath, FileMode.CreateNew, FileAccess.Write)) {
                file.SetLength(PdfLoadOptions.Default.Limits.MaxInputBytes + 1);
            }

            Action<string>[] routes = {
                input => PdfFormFiller.FillFieldsToBytes(input, EmptyFields),
                input => PdfFormFiller.FlattenFieldsToBytes(input),
                input => PdfFormFiller.FillAndFlattenFieldsToBytes(input, EmptyFields),
                input => PdfStamper.StampTextToBytes(input, "stamp"),
                input => PdfStamper.WatermarkTextToBytes(input, "watermark"),
                input => PdfStamper.StampImageToBytes(input, Array.Empty<byte>()),
                input => PdfStamper.WatermarkImageToBytes(input, Array.Empty<byte>())
            };
            foreach (Action<string> route in routes) {
                Assert.Equal(PdfReadLimitKind.InputBytes,
                    Assert.Throws<PdfReadLimitException>(() => route(inputPath)).Kind);
            }

            Assert.Throws<PdfReadLimitException>(() =>
                PdfFormFiller.FillFields(inputPath, formOutput, EmptyFields));
            Assert.Throws<PdfReadLimitException>(() =>
                PdfStamper.StampText(inputPath, stampOutput, "stamp"));
            Assert.False(File.Exists(formOutput));
            Assert.False(File.Exists(stampOutput));
        } finally {
            if (Directory.Exists(root)) Directory.Delete(root, recursive: true);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PageStampStreamsApplyTheImportedSourcesOwnInputLimit(bool underlay) {
        byte[] target = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Target"))
            .ToBytes();
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Imported source"))
            .ToBytes();
        using var targetStream = new MemoryStream(target);
        using var sourceStream = new ChunkedNonSeekableStream(source, maximumChunkSize: 3);
        var options = new PdfPageOverlayOptions {
            SourceReadOptions = new PdfLoadOptions {
                Limits = new PdfReadLimits { MaxInputBytes = 16 }
            }
        };

        PdfReadLimitException error = Assert.Throws<PdfReadLimitException>(() => underlay
            ? PdfStamper.UnderlayPage(targetStream, sourceStream, options)
            : PdfStamper.OverlayPage(targetStream, sourceStream, options));

        Assert.Equal(PdfReadLimitKind.InputBytes, error.Kind);
        Assert.InRange(sourceStream.BytesRead, 17, 19);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ImageStampStreamsApplyTheEncodedImageInputLimit(bool watermark) {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Image target"))
            .ToBytes();
        byte[] image = PdfPngTestImages.CreateRgbPng(25, 50, 75);
        using var imageStream = new ChunkedNonSeekableStream(image, maximumChunkSize: 3);
        var options = new PdfImageStampOptions { MaximumEncodedImageBytes = 16 };

        Assert.Throws<InvalidDataException>(() => watermark
            ? PdfStamper.WatermarkImage(pdf, imageStream, options)
            : PdfStamper.StampImage(pdf, imageStream, options));

        Assert.InRange(imageStream.BytesRead, 17, 19);
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
        private readonly int _maximumChunkSize;
        private int _position;

        internal ChunkedNonSeekableStream(byte[] bytes, int maximumChunkSize) {
            _bytes = bytes;
            _maximumChunkSize = maximumChunkSize;
        }

        internal int BytesRead => _position;
        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => _position; set => throw new NotSupportedException(); }
        public override int Read(byte[] buffer, int offset, int count) {
            int available = _bytes.Length - _position;
            if (available <= 0) return 0;
            int read = Math.Min(Math.Min(count, _maximumChunkSize), available);
            Buffer.BlockCopy(_bytes, _position, buffer, offset, read);
            _position += read;
            return read;
        }
        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }
}
