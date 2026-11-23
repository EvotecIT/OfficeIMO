using OfficeIMO.Reader;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public sealed class ReaderInputLimitTests {
    [Fact]
    public void LargeSnapshotBudgetsUseDeleteOnCloseTemporaryStorage() {
        byte[] payload = System.Text.Encoding.UTF8.GetBytes(
            "bounded temporary snapshot");
        // A forward-only source has no known length, so the large budget alone selects file storage.
        using var source = new ForwardOnlyStream(payload);
        Stream prepared = ReaderInputLimits.EnsureSeekableReadStream(
            source,
            maxInputBytes: 64L * 1024 * 1024 + 1,
            CancellationToken.None,
            out bool ownsStream);
        string snapshotPath = Assert.IsAssignableFrom<FileStream>(prepared).Name;
        string snapshotDirectory = Path.GetDirectoryName(snapshotPath)!;
        using (prepared) {
            Assert.True(ownsStream);
            Assert.True(ReaderInputLimits.IsSnapshotStream(prepared));
            Assert.NotEqual(Path.GetTempPath().TrimEnd(Path.DirectorySeparatorChar), snapshotDirectory.TrimEnd(Path.DirectorySeparatorChar));
            Assert.Throws<IOException>(() => new FileStream(snapshotPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite));
#if NET6_0_OR_GREATER
            if (!OperatingSystem.IsWindows()) {
                Assert.Equal(
                    UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute,
                    File.GetUnixFileMode(snapshotDirectory));
            }
#endif
            using var copy = new MemoryStream();
            prepared.CopyTo(copy);
            Assert.Equal(payload, copy.ToArray());

            Stream reused = ReaderInputLimits.EnsureSeekableReadStream(
                prepared,
                maxInputBytes: 64L * 1024 * 1024 + 1,
                CancellationToken.None,
                out bool ownsReused);
            Assert.Same(prepared, reused);
            Assert.False(ownsReused);
            Assert.Equal(0, reused.Position);
        }

        Assert.False(Directory.Exists(snapshotDirectory));
    }

    [Fact]
    public async Task SmallSeekableInputsSnapshotInMemoryWhateverTheBudget() {
        byte[] payload = System.Text.Encoding.UTF8.GetBytes("small seekable input");
        using var source = new MemoryStream(payload, writable: false);
        source.Position = 5;

        using (Stream prepared = ReaderInputLimits.EnsureSeekableReadStream(source, maxInputBytes: null, CancellationToken.None, out bool ownsStream)) {
            Assert.True(ownsStream);
            Assert.True(ReaderInputLimits.IsSnapshotStream(prepared));
            Assert.IsNotAssignableFrom<FileStream>(prepared);
            Assert.Equal(payload, ((MemoryStream)prepared).ToArray());
            Assert.Equal(5, source.Position);
        }

        using Stream preparedAsync = await ReaderInputLimits.EnsureSeekableReadStreamAsync(source, 64L * 1024 * 1024 + 1);
        Assert.IsNotAssignableFrom<FileStream>(preparedAsync);
        Assert.Equal(payload, ((MemoryStream)preparedAsync).ToArray());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task PublicSnapshotDoesNotTransferItsBuffer(bool asynchronous) {
        byte[] payload = { 0x15, 0x26, 0x37 };
        using var source = new MemoryStream(payload, writable: false);
        using Stream privateSnapshot = ReaderInputLimits.EnsureSeekableReadStream(
            source, payload.Length, inputLimitProbe: null, CancellationToken.None, out _);
        Assert.True(ReaderInputLimits.TryGetOwnedSnapshotMemory(privateSnapshot, out ReadOnlyMemory<byte> owned));
        Assert.Equal(payload, owned.ToArray());

        using Stream publicSnapshot = asynchronous
            ? await ReaderInputLimits.EnsureSeekableReadStreamAsync(source, payload.Length)
            : ReaderInputLimits.EnsureSeekableReadStream(source, payload.Length, CancellationToken.None, out _);
        Assert.False(ReaderInputLimits.TryGetOwnedSnapshotMemory(publicSnapshot, out _));
    }

    [Theory]
    [InlineData("chunks")]
    [InlineData("document")]
    [InlineData("async document")]
    [InlineData("container probe")]
    [InlineData("extension validator")]
    public async Task RegisteredStreamCallbackDoesNotReceiveATransferableSnapshot(string callback) {
        using var source = new MemoryStream(new byte[] { 0x15, 0x26, 0x37 }, writable: false);
        using Stream snapshot = ReaderInputLimits.EnsureSeekableReadStream(
            source, 3, inputLimitProbe: null, CancellationToken.None, out _);
        Assert.True(ReaderInputLimits.TryGetOwnedSnapshotMemory(snapshot, out _));

        bool invoked = false;
        void Observe(Stream received) {
            Assert.Same(snapshot, received);
            Assert.False(ReaderInputLimits.TryGetOwnedSnapshotMemory(received, out _));
            invoked = true;
        }

        var handler = ReaderHandlerDescriptor.Create(new ReaderHandlerRegistration {
            Id = "tests.snapshot-callback",
            Origin = ReaderHandlerOrigin.OfficeIMO,
            ReadStream = (input, _, _, _) => { Observe(input); return Array.Empty<ReaderChunk>(); },
            ReadDocumentStream = (input, _, _, _) => { Observe(input); return new OfficeDocumentReadResult(); },
            ReadDocumentStreamAsync = (input, _, _, _) => { Observe(input); return Task.FromResult(new OfficeDocumentReadResult()); },
            ProbeStream = (input, _, _, _) => { Observe(input); return false; },
            ExtensionValidationProbeStream = (input, _, _, _) => { Observe(input); return false; }
        });

        var options = new ReaderOptions();
        switch (callback) {
            case "chunks": handler.ReadStream!(snapshot, null, options, CancellationToken.None); break;
            case "document": handler.ReadDocumentStream!(snapshot, null, options, CancellationToken.None); break;
            case "async document": await handler.ReadDocumentStreamAsync!(snapshot, null, options, CancellationToken.None); break;
            case "container probe": handler.ProbeStream!(snapshot, null, options, CancellationToken.None); break;
            case "extension validator": handler.ExtensionValidationProbeStream!(snapshot, null, options, CancellationToken.None); break;
            default: throw new ArgumentOutOfRangeException(nameof(callback));
        }

        Assert.True(invoked);
        Assert.False(ReaderInputLimits.TryGetOwnedSnapshotMemory(snapshot, out _));
    }

    private sealed class ForwardOnlyStream : Stream {
        private readonly MemoryStream _inner;

        internal ForwardOnlyStream(byte[] bytes) => _inner = new MemoryStream(bytes, writable: false);

        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void Flush() { }
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        protected override void Dispose(bool disposing) {
            if (disposing) _inner.Dispose();
            base.Dispose(disposing);
        }
    }

    [Fact]
    public async Task LargeSnapshotAsyncDisposalRemovesPrivateDirectory() {
#if NET6_0_OR_GREATER
        byte[] payload = System.Text.Encoding.UTF8.GetBytes(
            "bounded asynchronous snapshot");
        using var source = new ForwardOnlyStream(payload);
        Stream prepared = await ReaderInputLimits.EnsureSeekableReadStreamAsync(
            source,
            maxInputBytes: 64L * 1024 * 1024 + 1,
            CancellationToken.None);
        string snapshotPath = Assert.IsAssignableFrom<FileStream>(prepared).Name;
        string snapshotDirectory = Path.GetDirectoryName(snapshotPath)!;

        await prepared.DisposeAsync();

        Assert.False(Directory.Exists(snapshotDirectory));
#else
        await Task.CompletedTask;
#endif
    }
}
