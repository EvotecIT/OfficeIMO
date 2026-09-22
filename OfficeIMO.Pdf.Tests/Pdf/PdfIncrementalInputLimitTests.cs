using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfIncrementalInputLimitTests {
    [Fact]
    public void IncrementalStreamEntrypointsRejectOversizedRemainingInputBeforeReading() {
        Action<Stream>[] routes = {
            input => PdfIncrementalUpdater.AnalyzeAppendOnlyMutation(input),
            input => PdfIncrementalUpdater.UpdateMetadata(input, title: "bounded"),
            input => PdfIncrementalUpdater.UpdateFormFields(input, new Dictionary<string, string>()),
            input => PdfMutationPlanner.Plan(input, PdfMutationOperation.UpdateMetadata)
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
    public void IncrementalStreamEntrypointsHonorCustomLimitsForNonSeekableInputs() {
        var options = new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxInputBytes = 1024 }
        };
        byte[] bytes = new byte[4096];

        using (var input = new ChunkedNonSeekableStream(bytes, 256)) {
            PdfReadLimitException error = Assert.Throws<PdfReadLimitException>(() =>
                PdfIncrementalUpdater.UpdateMetadata(input, title: "bounded", readOptions: options));
            Assert.Equal(PdfReadLimitKind.InputBytes, error.Kind);
            Assert.InRange(input.BytesRead, 1025, 1280);
        }

        using (var input = new ChunkedNonSeekableStream(bytes, 256)) {
            PdfReadLimitException error = Assert.Throws<PdfReadLimitException>(() =>
                PdfIncrementalUpdater.UpdateFormFields(
                    input,
                    new Dictionary<string, string>(),
                    options: null,
                    readOptions: options));
            Assert.Equal(PdfReadLimitKind.InputBytes, error.Kind);
            Assert.InRange(input.BytesRead, 1025, 1280);
        }
    }

    [Fact]
    public void IncrementalPathEntrypointsRejectOversizedInputWithoutChangingOutputs() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-incremental-limit-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(root, "oversized.pdf");
        string existingOutput = Path.Combine(root, "existing.pdf");
        string absentOutputDirectory = Path.Combine(root, "outputs");
        byte[] sentinel = { 8, 2, 8, 2 };
        try {
            Directory.CreateDirectory(root);
            using (var file = new FileStream(inputPath, FileMode.CreateNew, FileAccess.Write)) {
                file.SetLength(PdfLoadOptions.Default.Limits.MaxInputBytes + 1);
            }
            File.WriteAllBytes(existingOutput, sentinel);

            AssertInputLimit(() => PdfIncrementalUpdater.AnalyzeAppendOnlyMutation(inputPath));
            AssertInputLimit(() => PdfMutationPlanner.Plan(inputPath, PdfMutationOperation.UpdateMetadata));

            AssertInputLimit(() => PdfIncrementalUpdater.UpdateMetadata(inputPath, existingOutput, title: "bounded"));
            Assert.Equal(sentinel, File.ReadAllBytes(existingOutput));

            AssertInputLimit(() => PdfIncrementalUpdater.UpdateFormFields(
                inputPath,
                existingOutput,
                new Dictionary<string, string>()));
            Assert.Equal(sentinel, File.ReadAllBytes(existingOutput));

            AssertInputLimit(() => PdfIncrementalUpdater.ApplyExternalSignature(
                inputPath,
                Path.Combine(absentOutputDirectory, "signed.pdf"),
                new byte[] { 0x30, 0x01, 0x00 }));
            Assert.False(Directory.Exists(absentOutputDirectory));
        } finally {
            if (Directory.Exists(root)) Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void IncrementalStreamEntrypointsContinueFromCallerPosition() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("incremental source"))
            .ToBytes();
        byte[] form = PdfDocument.Create()
            .TextField("Name", value: "Ada")
            .ToBytes();

        using (var input = CreatePrefixedStream(source)) {
            PdfAppendOnlyMutationReport report = PdfIncrementalUpdater.AnalyzeAppendOnlyMutation(input);
            Assert.True(report.CanAppendMetadata);
            Assert.Equal(input.Length, input.Position);
        }

        using (var input = CreatePrefixedStream(source)) {
            byte[] updated = PdfIncrementalUpdater.UpdateMetadata(input, title: "Updated");
            Assert.Equal("Updated", PdfInspector.Inspect(updated).Metadata.Title);
            Assert.Equal(input.Length, input.Position);
        }

        using (var input = CreatePrefixedStream(form)) {
            byte[] updated = PdfIncrementalUpdater.UpdateFormFields(input, new Dictionary<string, string> {
                ["Name"] = "Grace"
            });
            Assert.Equal("Grace", Assert.Single(PdfInspector.Inspect(updated).FormFields).Value);
            Assert.Equal(input.Length, input.Position);
        }

        using (var input = CreatePrefixedStream(source)) {
            PdfMutationPlan plan = PdfMutationPlanner.Plan(input, PdfMutationOperation.UpdateMetadata);
            Assert.True(plan.CanExecute);
            Assert.Equal(input.Length, input.Position);
        }
    }

    private static void AssertInputLimit(Action action) {
        PdfReadLimitException error = Assert.Throws<PdfReadLimitException>(action);
        Assert.Equal(PdfReadLimitKind.InputBytes, error.Kind);
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
