using System.Security.Cryptography;
using System.Threading.Tasks;
using OfficeIMO.Reader;
using OfficeIMO.Reader.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Reader_capture_restores_caller_stream_when_cancelled_after_reading_starts() {
        using var package = CreatePagesPackage(includeBody: true, textBox: null, includePreview: false,
            bodyText: "Content must not be published after cancellation");
        using var cancellation = new System.Threading.CancellationTokenSource();
        using var stream = new CancelCaptureReadStream(package.ToArray(), cancellation);
        stream.Position = 5;

        Assert.False(cancellation.IsCancellationRequested);
        Assert.Throws<OperationCanceledException>(() => IWorkReaderAdapter.ReadDocument(stream,
            "cancelled.pages", new ReaderOptions(), new ReaderIWorkOptions(), cancellation.Token));

        Assert.True(stream.ReadStarted);
        Assert.True(cancellation.IsCancellationRequested);
        Assert.Equal(5, stream.Position);
        Assert.True(stream.CanRead);
        Assert.Equal(package.ToArray()[5], stream.ReadByte());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Reader_capture_keeps_iWork_package_limits_and_disabled_hashes(bool pathInput) {
        using var package = CreatePagesPackage(includeBody: true, textBox: null, includePreview: false,
            bodyText: "Bounded captured content");
        string root = Path.Combine(Path.GetTempPath(), "officeimo-iwork-capture-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string path = Path.Combine(root, "bounded.pages");
        try {
            File.WriteAllBytes(path, package.ToArray());
            var readerOptions = new ReaderOptions { ComputeHashes = false };
            var bounded = new ReaderIWorkOptions {
                ReadOptions = new OfficeIMO.IWork.IWorkReadOptions { MaximumPackageBytes = package.Length - 1 }
            };
            package.Position = 5;
            Assert.Throws<IOException>(() => pathInput
                ? IWorkReaderAdapter.ReadDocument(path, readerOptions, bounded, default)
                : IWorkReaderAdapter.ReadDocument(package, "bounded.pages", readerOptions, bounded, default));
            Assert.Equal(5, package.Position);
            var reader = new OfficeDocumentReaderBuilder().AddIWorkHandler().Build();
            OfficeDocumentReadResult read = pathInput ? reader.ReadDocument(path, readerOptions)
                : reader.ReadDocument(package, "bounded.pages", readerOptions);
            Assert.Equal(package.Length, read.Source.LengthBytes);
            Assert.Null(read.Source.SourceHash);
            Assert.Contains(read.Chunks, chunk => chunk.Text == "Bounded captured content");
            Assert.All(read.Chunks, chunk => { Assert.Null(chunk.SourceHash); Assert.Null(chunk.ChunkHash); });
            Assert.Equal(5, package.Position);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task Reader_rejects_path_replacement_while_iWork_adapter_identity_describes_captured_bytes(bool asynchronous, bool chunksOnly) {
        using var original = CreatePagesPackage(includeBody: true, textBox: null, includePreview: false,
            bodyText: "Captured original content");
        using var replacement = CreatePagesPackage(includeBody: true, textBox: null, includePreview: false,
            bodyText: "Replacement content of a different length");
        byte[] originalBytes = original.ToArray();
        string root = Path.Combine(Path.GetTempPath(), "officeimo-iwork-capture-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string path = Path.Combine(root, "mutable.pages");
        try {
            File.WriteAllBytes(path, originalBytes);
            OfficeDocumentReadResult? capturedResult = null;
            OfficeDocumentReadResult CaptureAndReplace(string input, ReaderOptions options, System.Threading.CancellationToken token) {
                if (chunksOnly) File.WriteAllBytes(input, replacement.ToArray());
                var captured = IWorkReaderAdapter.ReadDocument(input, options, new ReaderIWorkOptions(), token);
                capturedResult = captured;
                if (!chunksOnly) File.WriteAllBytes(input, replacement.ToArray());
                return captured;
            }
            var registration = new ReaderHandlerRegistration {
                Id = "officeimo.tests.iwork-capture", Kind = ReaderInputKind.IWork, Extensions = new[] { ".pages" },
                ReadPath = (input, options, token) => CaptureAndReplace(input, options, token).Chunks
            };
            if (!chunksOnly) registration.ReadDocumentPath = CaptureAndReplace;
            var reader = new OfficeDocumentReaderBuilder().AddHandler(registration).Build();
            var readOptions = new ReaderOptions { ComputeHashes = true };
            // Reader requires a stable path even when an adapter safely captures its own input.
            if (asynchronous) {
                await Assert.ThrowsAsync<IOException>(() => reader.ReadDocumentAsync(path, readOptions));
            } else {
                Assert.Throws<IOException>(() => reader.ReadDocument(path, readOptions));
            }
            OfficeDocumentReadResult result = Assert.IsType<OfficeDocumentReadResult>(capturedResult);
            byte[] capturedBytes = chunksOnly ? replacement.ToArray() : originalBytes;
            string expectedHash = Convert.ToHexString(SHA256.HashData(capturedBytes)).ToLowerInvariant();
            Assert.Equal(expectedHash, result.Source.SourceHash);
            Assert.Equal(capturedBytes.LongLength, result.Source.LengthBytes);
            Assert.All(result.Chunks, chunk => {
                Assert.Equal(expectedHash, chunk.SourceHash);
                Assert.Equal(capturedBytes.LongLength, chunk.SourceLengthBytes);
            });
            string capturedText = chunksOnly ? "Replacement content" : "Captured original content";
            string otherText = chunksOnly ? "Captured original content" : "Replacement content";
            Assert.Contains(result.Chunks, chunk => chunk.Text.Contains(capturedText));
            Assert.DoesNotContain(result.Chunks, chunk => chunk.Text.Contains(otherText));
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    private sealed class CancelCaptureReadStream(byte[] bytes, System.Threading.CancellationTokenSource cancellation)
        : MemoryStream(bytes, writable: false) {
        public bool ReadStarted { get; private set; }

        public override int Read(byte[] buffer, int offset, int count) {
            int read = base.Read(buffer, offset, count);
            if (read > 0) {
                ReadStarted = true;
                cancellation.Cancel();
            }
            return read;
        }
    }
}
