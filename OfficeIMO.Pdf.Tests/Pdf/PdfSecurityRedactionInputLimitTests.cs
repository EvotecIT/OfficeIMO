using OfficeIMO.Pdf;
using System.Threading;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfSecurityRedactionInputLimitTests {
    private static readonly PdfRedactionArea[] NoAreas = Array.Empty<PdfRedactionArea>();

    [Fact]
    public void DocumentUtilityStreamsRejectOversizedRemainingInputBeforeReading() {
        Action<Stream>[] routes = {
            input => PdfSanitizer.Sanitize(input),
            input => PdfOptimizer.Optimize(input),
            input => PdfRedactionPlanner.Plan(input, NoAreas),
            input => PdfRedactionApplier.Apply(input, NoAreas)
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
    public void DocumentUtilityPathsRejectOversizedInputWithoutChangingOutputs() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-security-redaction-limit-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(root, "input.pdf");
        string existingOutput = Path.Combine(root, "existing.pdf");
        string outputDirectory = Path.Combine(root, "outputs");
        byte[] sentinel = { 4, 2, 4, 2 };
        var encryption = new PdfStandardEncryptionOptions("open") { OwnerPassword = "owner" };
        try {
            Directory.CreateDirectory(root);
            using (var file = new FileStream(inputPath, FileMode.CreateNew, FileAccess.Write)) {
                file.SetLength(PdfLoadOptions.Default.Limits.MaxInputBytes + 1);
            }
            File.WriteAllBytes(existingOutput, sentinel);

            Action[] resultRoutes = {
                () => PdfSanitizer.Sanitize(inputPath),
                () => PdfOptimizer.Optimize(inputPath),
                () => PdfRedactionPlanner.Plan(inputPath, NoAreas),
                () => PdfRedactionApplier.ApplyToBytes(inputPath, NoAreas)
            };
            foreach (Action route in resultRoutes) {
                Assert.Equal(PdfReadLimitKind.InputBytes,
                    Assert.Throws<PdfReadLimitException>(route).Kind);
            }

            Action[] existingOutputRoutes = {
                () => PdfSecurityEditor.Encrypt(inputPath, existingOutput, encryption),
                () => PdfSecurityEditor.Decrypt(inputPath, existingOutput, "owner"),
                () => PdfOptimizer.Optimize(inputPath, existingOutput),
                () => PdfRedactionApplier.Apply(inputPath, existingOutput, NoAreas)
            };
            foreach (Action route in existingOutputRoutes) {
                Assert.Equal(PdfReadLimitKind.InputBytes,
                    Assert.Throws<PdfReadLimitException>(route).Kind);
                Assert.Equal(sentinel, File.ReadAllBytes(existingOutput));
            }

            Action[] absentOutputRoutes = {
                () => PdfSecurityEditor.Encrypt(inputPath, Path.Combine(outputDirectory, "encrypted.pdf"), encryption),
                () => PdfOptimizer.Optimize(inputPath, Path.Combine(outputDirectory, "optimized.pdf")),
                () => PdfRedactionApplier.Apply(inputPath, Path.Combine(outputDirectory, "redacted.pdf"), NoAreas)
            };
            foreach (Action route in absentOutputRoutes) {
                Assert.Equal(PdfReadLimitKind.InputBytes,
                    Assert.Throws<PdfReadLimitException>(route).Kind);
                Assert.False(Directory.Exists(outputDirectory));
            }
        } finally {
            if (Directory.Exists(root)) Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void RedactionPathsHonorExplicitSourceInputLimit() {
        string path = Path.Combine(Path.GetTempPath(), "officeimo-redaction-limit-" + Guid.NewGuid().ToString("N") + ".pdf");
        try {
            File.WriteAllBytes(path, new byte[5]);
            var options = new PdfLoadOptions { Limits = new PdfReadLimits { MaxInputBytes = 4 } };

            Assert.Equal(PdfReadLimitKind.InputBytes,
                Assert.Throws<PdfReadLimitException>(() => PdfRedactionPlanner.Plan(path, NoAreas, options: options)).Kind);
            Assert.Equal(PdfReadLimitKind.InputBytes,
                Assert.Throws<PdfReadLimitException>(() => PdfRedactionApplier.ApplyToBytes(path, NoAreas, readOptions: options)).Kind);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void BoundedSourceCancellationAndNonseekableLimitsApplyDuringSnapshot() {
        var options = new PdfLoadOptions { Limits = new PdfReadLimits { MaxInputBytes = 4 } };
        using var cancelled = new TrackingNonSeekableStream(new byte[3]);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() => {
            _ = PdfDocumentSource.FromRemainingStream(cancelled, options, cancellation.Token);
        });
        Assert.False(cancelled.WasRead);

        using var oversized = new TrackingNonSeekableStream(new byte[5]);
        PdfReadLimitException error = Assert.Throws<PdfReadLimitException>(() =>
            PdfDocumentSource.FromRemainingStream(oversized, options, CancellationToken.None));
        Assert.Equal(PdfReadLimitKind.InputBytes, error.Kind);
        Assert.True(oversized.WasRead);
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

    private sealed class TrackingNonSeekableStream : Stream {
        private readonly byte[] _bytes;
        private int _offset;

        internal TrackingNonSeekableStream(byte[] bytes) => _bytes = bytes;
        internal bool WasRead { get; private set; }
        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override int Read(byte[] buffer, int offset, int count) {
            WasRead = true;
            int available = Math.Min(count, _bytes.Length - _offset);
            if (available == 0) return 0;
            Buffer.BlockCopy(_bytes, _offset, buffer, offset, available);
            _offset += available;
            return available;
        }
        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }
}
