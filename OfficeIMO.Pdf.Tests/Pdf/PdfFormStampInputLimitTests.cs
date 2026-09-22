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
