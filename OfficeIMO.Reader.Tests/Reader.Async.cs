using OfficeIMO.PowerPoint;
using OfficeIMO.Reader;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderAsyncTests {
    [Fact]
    public async Task OfficeDocumentReader_UsesNativeAsyncPathHandlerAndAdvertisesCapability() {
        const string extension = ".asyncpathix";
        string file = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + extension);
        File.WriteAllText(file, "input");

        try {
            OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
                .AddHandler(new ReaderHandlerRegistration {
                    Id = "officeimo.tests.async.path",
                    Kind = ReaderInputKind.Text,
                    Extensions = new[] { extension },
                    ReadDocumentPathAsync = async (path, options, cancellationToken) => {
                        await Task.Yield();
                        cancellationToken.ThrowIfCancellationRequested();
                        return CreateResult(path, "native-async-path");
                    }
                })
                .Build();

            OfficeDocumentReadResult result = await reader.ReadDocumentAsync(file);
            IReadOnlyList<ReaderChunk> chunks = await reader.ReadAsync(file);
            ReaderHandlerCapability capability = Assert.Single(
                reader.GetCapabilities(),
                item => item.Id == "officeimo.tests.async.path");

            Assert.Equal("native-async-path", Assert.Single(result.Chunks).Text);
            Assert.Equal("native-async-path", Assert.Single(chunks).Text);
            Assert.False(string.IsNullOrWhiteSpace(chunks[0].SourceId));
            Assert.True(capability.SupportsPath);
            Assert.True(capability.SupportsDocumentPath);
            Assert.True(capability.SupportsAsyncPath);
            Assert.False(capability.SupportsAsyncStream);
            Assert.Throws<InvalidOperationException>(() => reader.ReadDocument(file));
        } finally {
            File.Delete(file);
        }
    }

    [Fact]
    public async Task OfficeDocumentReader_UsesBoundedFallbackForSynchronousPathHandler() {
        const string extension = ".syncpathix";
        string file = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + extension);
        File.WriteAllText(file, "input");

        try {
            OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
                .WithMaxConcurrentReads(1)
                .AddHandler(new ReaderHandlerRegistration {
                    Id = "officeimo.tests.sync.path",
                    Kind = ReaderInputKind.Text,
                    Extensions = new[] { extension },
                    ReadDocumentPath = (path, options, cancellationToken) => {
                        cancellationToken.ThrowIfCancellationRequested();
                        return CreateResult(path, "bounded-sync-path");
                    }
                })
                .Build();

            OfficeDocumentReadResult result = await reader.ReadDocumentAsync(file);
            IReadOnlyList<ReaderChunk> chunks = await reader.ReadAsync(file);
            ReaderHandlerCapability capability = Assert.Single(
                reader.GetCapabilities(),
                item => item.Id == "officeimo.tests.sync.path");

            Assert.Equal("bounded-sync-path", Assert.Single(result.Chunks).Text);
            Assert.Equal("bounded-sync-path", Assert.Single(chunks).Text);
            Assert.True(capability.SupportsPath);
            Assert.False(capability.SupportsAsyncPath);
            Assert.Equal(1, reader.MaxConcurrentReads);
        } finally {
            File.Delete(file);
        }
    }

    [Fact]
    public async Task OfficeDocumentReader_UsesAsyncBoundedSnapshotForNonSeekableStream() {
        const string extension = ".asyncstreamix";
        bool handlerInvoked = false;
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddHandler(new ReaderHandlerRegistration {
                Id = "officeimo.tests.async.stream",
                Kind = ReaderInputKind.Text,
                Extensions = new[] { extension },
                ReadDocumentStreamAsync = async (stream, sourceName, options, cancellationToken) => {
                    handlerInvoked = true;
                    Assert.True(stream.CanSeek);
                    var buffer = new byte[64];
                    int bytesRead = await stream.ReadAsync(buffer, 0, buffer.Length, cancellationToken);
                    return CreateResult(sourceName ?? "memory", "bytes:" + bytesRead);
                }
            })
            .Build();
        using var stream = new AsyncOnlyNonSeekableStream(new byte[32]);

        OfficeDocumentReadResult result = await reader.ReadDocumentAsync(
            stream,
            "sample" + extension,
            new ReaderOptions { MaxInputBytes = 64 });

        Assert.True(handlerInvoked);
        Assert.Equal("bytes:32", Assert.Single(result.Chunks).Text);
    }

    [Fact]
    public async Task OfficeDocumentReader_RejectsOversizedAsyncStreamBeforeHandlerDispatch() {
        const string extension = ".asynclimitix";
        bool handlerInvoked = false;
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddHandler(new ReaderHandlerRegistration {
                Id = "officeimo.tests.async.limit",
                Kind = ReaderInputKind.Text,
                Extensions = new[] { extension },
                ReadDocumentStreamAsync = (stream, sourceName, options, cancellationToken) => {
                    handlerInvoked = true;
                    return Task.FromResult(CreateResult(sourceName ?? "memory", "unexpected"));
                }
            })
            .Build();
        using var stream = new AsyncOnlyNonSeekableStream(new byte[128]);

        IOException exception = await Assert.ThrowsAsync<IOException>(() => reader.ReadDocumentAsync(
            stream,
            "sample" + extension,
            new ReaderOptions { MaxInputBytes = 16 }));

        Assert.Contains("Input exceeds MaxInputBytes", exception.Message, StringComparison.Ordinal);
        Assert.False(handlerInvoked);
    }

    [Fact]
    public async Task OfficeDocumentReader_ReadDocumentsAsync_BoundsConcurrencyAndPreservesOrder() {
        const string extension = ".asyncbatchix";
        var files = Enumerable.Range(0, 5)
            .Select(index => Path.Combine(Path.GetTempPath(), $"{index}-{Guid.NewGuid():N}{extension}"))
            .ToArray();
        foreach (string file in files) File.WriteAllText(file, "input");

        int active = 0;
        int maximumActive = 0;
        var twoStarted = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        var release = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);

        try {
            OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
                .WithMaxConcurrentReads(3)
                .AddHandler(new ReaderHandlerRegistration {
                    Id = "officeimo.tests.async.batch",
                    Kind = ReaderInputKind.Text,
                    Extensions = new[] { extension },
                    ReadDocumentPathAsync = async (path, options, cancellationToken) => {
                        int current = Interlocked.Increment(ref active);
                        UpdateMaximum(ref maximumActive, current);
                        if (current == 2) twoStarted.TrySetResult(true);
                        try {
                            await release.Task.ConfigureAwait(false);
                            cancellationToken.ThrowIfCancellationRequested();
                            return CreateResult(path, Path.GetFileName(path));
                        } finally {
                            Interlocked.Decrement(ref active);
                        }
                    }
                })
                .Build();

            Task<IReadOnlyList<OfficeDocumentReadResult>> batch = reader.ReadDocumentsAsync(
                files,
                batchOptions: new ReaderBatchOptions {
                    MaxDegreeOfParallelism = 2,
                    MaxDocuments = 5
                });

            Task started = await Task.WhenAny(twoStarted.Task, Task.Delay(TimeSpan.FromSeconds(5)));
            Assert.Same(twoStarted.Task, started);
            Assert.Equal(2, maximumActive);
            release.TrySetResult(true);

            IReadOnlyList<OfficeDocumentReadResult> results = await batch;
            Assert.Equal(files, results.Select(result => result.Source.Path).ToArray());
            Assert.Equal(2, maximumActive);
            Assert.Equal(3, reader.MaxConcurrentReads);
        } finally {
            release.TrySetResult(true);
            foreach (string file in files) File.Delete(file);
        }
    }

    [Fact]
    public async Task OfficeDocumentReader_ReadDocumentsAsync_EnforcesMaxDocumentsBeforeReading() {
        OfficeDocumentReader reader = OfficeDocumentReader.Default;

        InvalidOperationException exception = await Assert.ThrowsAsync<InvalidOperationException>(() => reader.ReadDocumentsAsync(
            new[] { "first.docx", "second.docx" },
            batchOptions: new ReaderBatchOptions { MaxDocuments = 1 }));

        Assert.Contains("MaxDocuments", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public async Task DocumentReader_ReadAsync_UsesWorkerFallbackForBuiltInReader() {
        string file = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".txt");
        File.WriteAllText(file, "built-in async fallback");

        try {
            IReadOnlyList<ReaderChunk> chunks = await OfficeDocumentReader.Default.ReadAsync(file);

            Assert.Contains(chunks, chunk => chunk.Text.Contains("built-in async fallback", StringComparison.Ordinal));
        } finally {
            File.Delete(file);
        }
    }

    [Fact]
    public async Task DocumentReader_ReadDocumentAsync_CancelsEncryptedDetectionAfterInitialProbe() {
        const string password = "reader-async-cancel-pass";
        byte[] encrypted;
        using (PowerPointPresentation source =
               PowerPointPresentation.Create()) {
            source.AddSlide().AddTextBox("Encrypted cancellation probe");
            encrypted = source.ToEncryptedBytes(password);
        }
        using var cancellation = new CancellationTokenSource();
        using var stream = new CancelAfterFirstAsyncReadStream(encrypted,
            cancellation.Cancel);
        var options = new ReaderOptions {
            OpenPassword = password,
            DetectionMaxProbeBytes = 256
        };

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            OfficeDocumentReader.Default.ReadDocumentAsync(stream,
                sourceName: null, options, cancellation.Token));
    }

    [Fact]
    public void DocumentReader_ReadDocument_CancelsEncryptedDetectionAfterInitialProbe() {
        const string password = "reader-sync-cancel-pass";
        byte[] encrypted;
        using (PowerPointPresentation source =
               PowerPointPresentation.Create()) {
            source.AddSlide().AddTextBox("Encrypted sync cancellation probe");
            encrypted = source.ToEncryptedBytes(password);
        }
        using var cancellation = new CancellationTokenSource();
        using var stream = new CancelAfterFirstSyncReadStream(encrypted,
            cancellation.Cancel);
        var options = new ReaderOptions {
            OpenPassword = password,
            DetectionMaxProbeBytes = 256
        };

        Assert.ThrowsAny<OperationCanceledException>(() =>
            OfficeDocumentReader.Default.ReadDocument(stream,
                sourceName: null, options, cancellation.Token));
    }

    [Fact]
    public void DocumentReader_Read_CancelsSourceHashWithoutScanningRemainder() {
        byte[] bytes = new byte[256 * 1024];
        using var cancellation = new CancellationTokenSource();
        using var stream = new CancelAfterFirstSyncReadStream(bytes,
            cancellation.Cancel);

        Assert.ThrowsAny<OperationCanceledException>(() =>
            OfficeDocumentReader.Default.Read(stream, "large.txt",
                new ReaderOptions { ComputeHashes = true },
                cancellation.Token).ToArray());

        Assert.InRange(stream.BytesRead, 1, 81920);
    }

    private static OfficeDocumentReadResult CreateResult(string path, string text) {
        return new OfficeDocumentReadResult {
            Kind = ReaderInputKind.Text,
            Source = new OfficeDocumentSource { Path = path },
            Chunks = new[] {
                new ReaderChunk {
                    Id = "async:0001",
                    Kind = ReaderInputKind.Text,
                    Location = new ReaderLocation { Path = path },
                    Text = text
                }
            }
        };
    }

    private static void UpdateMaximum(ref int maximum, int candidate) {
        while (true) {
            int current = Volatile.Read(ref maximum);
            if (candidate <= current || Interlocked.CompareExchange(ref maximum, candidate, current) == current) {
                return;
            }
        }
    }

    private sealed class AsyncOnlyNonSeekableStream : Stream {
        private readonly MemoryStream _inner;

        public AsyncOnlyNonSeekableStream(byte[] bytes) {
            _inner = new MemoryStream(bytes, writable: false);
        }

        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        public override int Read(byte[] buffer, int offset, int count) {
            throw new InvalidOperationException("Synchronous reads are not allowed by this test stream.");
        }

        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) {
            return _inner.ReadAsync(buffer, offset, count, cancellationToken);
        }

        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        protected override void Dispose(bool disposing) {
            if (disposing) _inner.Dispose();
            base.Dispose(disposing);
        }
    }

    private sealed class CancelAfterFirstAsyncReadStream : Stream {
        private readonly MemoryStream _inner;
        private readonly Action _cancel;
        private int _asyncReads;

        public CancelAfterFirstAsyncReadStream(byte[] bytes,
            Action cancel) {
            _inner = new MemoryStream(bytes, writable: false);
            _cancel = cancel;
        }

        public override bool CanRead => true;
        public override bool CanSeek => true;
        public override bool CanWrite => false;
        public override long Length => _inner.Length;
        public override long Position {
            get => _inner.Position;
            set => _inner.Position = value;
        }

        public override int Read(byte[] buffer, int offset, int count) =>
            _inner.Read(buffer, offset, count);

        public override async Task<int> ReadAsync(byte[] buffer, int offset,
            int count, CancellationToken cancellationToken) {
            int read = await _inner.ReadAsync(buffer, offset, count,
                cancellationToken).ConfigureAwait(false);
            if (Interlocked.Increment(ref _asyncReads) == 1) {
                _cancel();
            }
            return read;
        }

        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) =>
            _inner.Seek(offset, origin);
        public override void SetLength(long value) =>
            throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) =>
            throw new NotSupportedException();

        protected override void Dispose(bool disposing) {
            if (disposing) _inner.Dispose();
            base.Dispose(disposing);
        }
    }

    private sealed class CancelAfterFirstSyncReadStream : Stream {
        private readonly MemoryStream _inner;
        private readonly Action _cancel;
        private int _reads;

        public CancelAfterFirstSyncReadStream(byte[] bytes, Action cancel) {
            _inner = new MemoryStream(bytes, writable: false);
            _cancel = cancel;
        }

        public override bool CanRead => true;
        public override bool CanSeek => true;
        public override bool CanWrite => false;
        public long BytesRead { get; private set; }
        public override long Length => _inner.Length;
        public override long Position {
            get => _inner.Position;
            set => _inner.Position = value;
        }

        public override int Read(byte[] buffer, int offset, int count) {
            int read = _inner.Read(buffer, offset, count);
            BytesRead += read;
            if (Interlocked.Increment(ref _reads) == 1) _cancel();
            return read;
        }

        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) =>
            _inner.Seek(offset, origin);
        public override void SetLength(long value) =>
            throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) =>
            throw new NotSupportedException();

        protected override void Dispose(bool disposing) {
            if (disposing) _inner.Dispose();
            base.Dispose(disposing);
        }
    }
}
