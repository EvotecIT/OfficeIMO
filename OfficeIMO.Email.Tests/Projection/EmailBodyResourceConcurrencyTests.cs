using OfficeIMO.Email;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class EmailBodyResourceConcurrencyTests {
    [Fact]
    public async Task Independent_resources_read_concurrently_through_the_sync_surface() {
        var blockedContent = new GateContentSource(1);
        var immediateContent = new ImmediateContentSource(2);
        EmailBodyProjectionResult projection = CreateProjection(
            maxResourceBytes: 2,
            maxTotalResourceBytes: 2,
            blockedContent,
            immediateContent);
        using Stream blocked = projection.Resources[0].OpenReadStream();
        using Stream immediate = projection.Resources[1].OpenReadStream();
        var blockedBuffer = new byte[1];
        var immediateBuffer = new byte[1];

        Task<int> blockedRead = Task.Run(() => blocked.Read(blockedBuffer, 0, 1));
        await blockedContent.ReadStarted;
        Task<int> immediateRead = Task.Run(() => immediate.Read(immediateBuffer, 0, 1));

        Assert.Same(immediateRead, await Task.WhenAny(immediateRead, Task.Delay(TimeSpan.FromSeconds(5))));
        Assert.Equal(1, await immediateRead);
        Assert.Equal(2, immediateBuffer[0]);

        blockedContent.ReleaseRead();
        Assert.Equal(1, await blockedRead);
        Assert.Equal(1, blockedBuffer[0]);
    }

    [Fact]
    public async Task Independent_resources_read_concurrently_while_aggregate_capacity_remains() {
        var blockedContent = new GateContentSource(1);
        var immediateContent = new ImmediateContentSource(2);
        EmailBodyProjectionResult projection = CreateProjection(
            maxResourceBytes: 2,
            maxTotalResourceBytes: 2,
            blockedContent,
            immediateContent);
        using Stream blocked = await projection.Resources[0].OpenReadStreamAsync();
        using Stream immediate = await projection.Resources[1].OpenReadStreamAsync();

        Task<int> blockedRead = blocked.ReadAsync(new byte[1], 0, 1);
        await blockedContent.ReadStarted;
        var immediateBuffer = new byte[1];
        Task<int> immediateRead = immediate.ReadAsync(immediateBuffer, 0, 1);

        Assert.Same(immediateRead, await Task.WhenAny(immediateRead, Task.Delay(TimeSpan.FromSeconds(5))));
        Assert.Equal(1, await immediateRead);
        Assert.Equal(2, immediateBuffer[0]);

        blockedContent.ReleaseRead();
        Assert.Equal(1, await blockedRead);
    }

    [Fact]
    public async Task Cancellation_while_waiting_for_an_aggregate_reservation_does_not_poison_the_stream() {
        var blockedContent = new GateContentSource(1);
        var immediateContent = new ImmediateContentSource(2);
        EmailBodyProjectionResult projection = CreateProjection(
            maxResourceBytes: 2,
            maxTotalResourceBytes: 2,
            blockedContent,
            immediateContent);
        using Stream blocked = await projection.Resources[0].OpenReadStreamAsync();
        using Stream waiting = await projection.Resources[1].OpenReadStreamAsync();

        Task<int> blockedRead = blocked.ReadAsync(new byte[2], 0, 2);
        await blockedContent.ReadStarted;
        using (var cancellation = new CancellationTokenSource(TimeSpan.FromMilliseconds(100))) {
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
                waiting.ReadAsync(new byte[1], 0, 1, cancellation.Token));
        }

        blockedContent.ReleaseRead();
        Assert.Equal(1, await blockedRead);
        var buffer = new byte[1];
        Assert.Equal(1, await waiting.ReadAsync(buffer, 0, 1));
        Assert.Equal(2, buffer[0]);
    }

    [Fact]
    public async Task Concurrent_stream_does_not_read_after_same_resource_is_poisoned() {
        var content = new PoisoningContentSource();
        EmailBodyProjectionResult projection = CreateProjection(
            maxResourceBytes: 3,
            maxTotalResourceBytes: 16,
            content);
        EmailBodyResource resource = projection.Resources[0];
        using Stream first = await resource.OpenReadStreamAsync();
        using Stream second = await resource.OpenReadStreamAsync();
        Assert.Equal(3, await first.ReadAsync(new byte[3], 0, 3));

        Task<int> poison = first.ReadAsync(new byte[1], 0, 1);
        await content.ProbeStarted;
        Task<int> waiting = second.ReadAsync(new byte[1], 0, 1);
        Assert.False(waiting.IsCompleted);

        content.ReleaseProbe();
        EmailLimitExceededException firstFailure = await Assert.ThrowsAsync<EmailLimitExceededException>(
            async () => await poison);
        EmailLimitExceededException secondFailure = await Assert.ThrowsAsync<EmailLimitExceededException>(
            async () => await waiting);

        Assert.Equal("EmailBodyProjectionOptions.MaxResourceBytes", firstFailure.LimitName);
        Assert.Equal("EmailBodyProjectionOptions.MaxResourceBytes", secondFailure.LimitName);
        Assert.Equal(0, content.SecondStreamReadCount);
    }

    private static EmailBodyProjectionResult CreateProjection(
        long maxResourceBytes,
        long maxTotalResourceBytes,
        params IEmailContentSource[] sources) {
        var document = new EmailDocument { Body = { Html = "<p>resources</p>" } };
        for (int index = 0; index < sources.Length; index++) {
            document.Attachments.Add(new EmailAttachment {
                ContentId = "resource-" + index,
                IsInline = true,
                ContentSource = sources[index],
                Length = 0
            });
        }
        return EmailBodyProjection.Create(document,
            new EmailBodyProjectionOptions {
                MaxResourceBytes = maxResourceBytes,
                MaxTotalResourceBytes = maxTotalResourceBytes
            });
    }

    private sealed class GateContentSource : IEmailContentSource {
        private readonly byte _value;
        private readonly TaskCompletionSource<bool> _readStarted =
            new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        private readonly TaskCompletionSource<bool> _releaseRead =
            new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);

        internal GateContentSource(byte value) {
            _value = value;
        }

        public long? Length => null;
        internal Task ReadStarted => _readStarted.Task;
        internal void ReleaseRead() => _releaseRead.TrySetResult(true);
        public Stream OpenRead() => new GateReadStream(_value, _readStarted, _releaseRead.Task);
        public Task<Stream> OpenReadAsync(CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            return Task.FromResult(OpenRead());
        }
    }

    private sealed class ImmediateContentSource : IEmailContentSource {
        private readonly byte _value;

        internal ImmediateContentSource(byte value) {
            _value = value;
        }

        public long? Length => null;
        public Stream OpenRead() => new MemoryStream(new[] { _value }, writable: false);
        public Task<Stream> OpenReadAsync(CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            return Task.FromResult(OpenRead());
        }
    }

    private sealed class PoisoningContentSource : IEmailContentSource {
        private readonly TaskCompletionSource<bool> _probeStarted =
            new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        private readonly TaskCompletionSource<bool> _releaseProbe =
            new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        private int _openCount;

        public long? Length => null;
        internal Task ProbeStarted => _probeStarted.Task;
        internal int SecondStreamReadCount { get; private set; }
        internal void ReleaseProbe() => _releaseProbe.TrySetResult(true);

        public Stream OpenRead() {
            int open = Interlocked.Increment(ref _openCount);
            return open == 1
                ? new PoisoningReadStream(_probeStarted, _releaseProbe.Task)
                : new CountingReadStream(() => SecondStreamReadCount++);
        }

        public Task<Stream> OpenReadAsync(CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            return Task.FromResult(OpenRead());
        }
    }

    private sealed class GateReadStream : Stream {
        private readonly byte _value;
        private readonly TaskCompletionSource<bool> _readStarted;
        private readonly Task _releaseRead;
        private bool _read;

        internal GateReadStream(byte value, TaskCompletionSource<bool> readStarted, Task releaseRead) {
            _value = value;
            _readStarted = readStarted;
            _releaseRead = releaseRead;
        }

        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override int Read(byte[] buffer, int offset, int count) {
            if (_read) return 0;
            _readStarted.TrySetResult(true);
            _releaseRead.GetAwaiter().GetResult();
            buffer[offset] = _value;
            _read = true;
            return 1;
        }
        public override async Task<int> ReadAsync(byte[] buffer, int offset, int count,
            CancellationToken cancellationToken) {
            if (_read) return 0;
            _readStarted.TrySetResult(true);
            await WaitWithCancellationAsync(_releaseRead, cancellationToken);
            buffer[offset] = _value;
            _read = true;
            return 1;
        }
        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }

    private sealed class PoisoningReadStream : Stream {
        private readonly MemoryStream _content = new MemoryStream(new byte[] { 1, 2, 3, 4 }, writable: false);
        private readonly TaskCompletionSource<bool> _probeStarted;
        private readonly Task _releaseProbe;

        internal PoisoningReadStream(TaskCompletionSource<bool> probeStarted, Task releaseProbe) {
            _probeStarted = probeStarted;
            _releaseProbe = releaseProbe;
        }

        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => _content.Length;
        public override long Position { get => _content.Position; set => throw new NotSupportedException(); }
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override async Task<int> ReadAsync(byte[] buffer, int offset, int count,
            CancellationToken cancellationToken) {
            if (_content.Position >= 3) {
                _probeStarted.TrySetResult(true);
                await WaitWithCancellationAsync(_releaseProbe, cancellationToken);
            }
            return _content.Read(buffer, offset, count);
        }
        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        protected override void Dispose(bool disposing) {
            if (disposing) _content.Dispose();
            base.Dispose(disposing);
        }
    }

    private sealed class CountingReadStream : Stream {
        private readonly Action _onRead;

        internal CountingReadStream(Action onRead) {
            _onRead = onRead;
        }

        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => 1;
        public override long Position { get => 0; set => throw new NotSupportedException(); }
        public override int Read(byte[] buffer, int offset, int count) {
            _onRead();
            buffer[offset] = 9;
            return 1;
        }
        public override Task<int> ReadAsync(byte[] buffer, int offset, int count,
            CancellationToken cancellationToken) {
            cancellationToken.ThrowIfCancellationRequested();
            return Task.FromResult(Read(buffer, offset, count));
        }
        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }

    private static async Task WaitWithCancellationAsync(Task operation, CancellationToken cancellationToken) {
        if (!cancellationToken.CanBeCanceled) {
            await operation;
            return;
        }
        var canceled = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        using (cancellationToken.Register(() => canceled.TrySetCanceled())) {
            await await Task.WhenAny(operation, canceled.Task);
        }
    }
}
