using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfAnnotationMetadataInputLimitTests {
    [Fact]
    public void AnnotationFlatteningAndMetadataStreamsRejectOversizedRemainingInputBeforeReading() {
        Action<Stream>[] routes = {
            input => PdfAnnotationFlattener.FlattenVisualAnnotations(input),
            input => PdfMetadataEditor.UpdateMetadata(input, title: "updated"),
            input => PdfMetadataEditor.ReplaceMetadata(input, new PdfMetadata { Title = "replacement" }),
            input => PdfMetadataEditor.SynchronizeMetadata(input, title: "synchronized")
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
    public void AnnotationAndMetadataPathsRejectOversizedInputWithoutCreatingOutputs() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-annotation-metadata-limit-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(root, "input.pdf");
        string outputDirectory = Path.Combine(root, "outputs");
        try {
            Directory.CreateDirectory(root);
            using (var file = new FileStream(inputPath, FileMode.CreateNew, FileAccess.Write)) {
                file.SetLength(PdfLoadOptions.Default.Limits.MaxInputBytes + 1);
            }

            Action[] byteRoutes = {
                () => PdfAnnotationFlattener.FlattenVisualAnnotationsToBytes(inputPath),
                () => PdfMetadataEditor.UpdateMetadataToBytes(inputPath, title: "updated"),
                () => PdfMetadataEditor.ReplaceMetadataToBytes(inputPath, new PdfMetadata { Title = "replacement" }),
                () => PdfMetadataEditor.SynchronizeMetadataToBytes(inputPath, title: "synchronized")
            };
            foreach (Action route in byteRoutes) {
                PdfReadLimitException error = Assert.Throws<PdfReadLimitException>(route);
                Assert.Equal(PdfReadLimitKind.InputBytes, error.Kind);
            }

            Action[] outputRoutes = {
                () => PdfAnnotationEditor.RemoveAnnotations(inputPath, Path.Combine(outputDirectory, "removed.pdf")),
                () => PdfAnnotationEditor.UpdateAnnotation(inputPath, Path.Combine(outputDirectory, "updated.pdf"), 1, new PdfAnnotationUpdateOptions()),
                () => PdfAnnotationFlattener.FlattenVisualAnnotations(inputPath, Path.Combine(outputDirectory, "flattened.pdf")),
                () => PdfMetadataEditor.UpdateMetadata(inputPath, Path.Combine(outputDirectory, "metadata.pdf"), title: "updated"),
                () => PdfMetadataEditor.ReplaceMetadata(inputPath, Path.Combine(outputDirectory, "replacement.pdf"), new PdfMetadata()),
                () => PdfMetadataEditor.SynchronizeMetadata(inputPath, Path.Combine(outputDirectory, "synchronized.pdf"), title: "updated")
            };
            foreach (Action route in outputRoutes) {
                PdfReadLimitException error = Assert.Throws<PdfReadLimitException>(route);
                Assert.Equal(PdfReadLimitKind.InputBytes, error.Kind);
                Assert.False(Directory.Exists(outputDirectory));
            }
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
