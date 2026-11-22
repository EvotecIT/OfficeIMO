using OfficeIMO.Core.Internal;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class OfficeStreamReaderContracts {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ForwardOnlyReadsConsumeAtMostOneByteBeyondTheLimit(bool asynchronous) {
        using var source = new CountingForwardStream(new byte[100_000]);
        InvalidDataException exception = asynchronous
            ? await Assert.ThrowsAsync<InvalidDataException>(() => OfficeStreamReader.ReadRemainingBytesAsync(source, maxBytes: 64))
            : Assert.Throws<InvalidDataException>(() => OfficeStreamReader.ReadRemainingBytes(source, maxBytes: 64));
        Assert.True(OfficeStreamReader.IsSizeLimitException(exception));
        Assert.Equal(65, source.BytesRead);
        Assert.True(source.CanRead);
    }

    [Theory]
    [InlineData(false, 64)]
    [InlineData(true, 64)]
    // Spans several of the growing chunks a forward read collects into, ending part-way through one.
    [InlineData(false, 81920 * 7 + 17)]
    [InlineData(true, 81920 * 7 + 17)]
    public async Task ExactLimitAndUnboundedMaximumReturnAllBytes(bool asynchronous, int length) {
        byte[] bytes = Enumerable.Range(0, length).Select(index => (byte)(index * 31 + index / 251)).ToArray();
        foreach (long limit in new[] { length, long.MaxValue }) {
            using var source = new CountingForwardStream(bytes);
            byte[] actual = asynchronous
                ? await OfficeStreamReader.ReadRemainingBytesAsync(source, maxBytes: limit)
                : OfficeStreamReader.ReadRemainingBytes(source, maxBytes: limit);
            Assert.Equal(bytes, actual);
            Assert.Equal(length, source.BytesRead);
            Assert.True(source.CanRead);
        }
    }

    [Fact]
    public void LegacyBufferKeepsItsLimitMessageAndForwardConsumption() {
        using var source = new CountingForwardStream(new byte[100_000]);
        var exception = Assert.Throws<InvalidDataException>(() =>
            OfficeLegacyImportBuffer.ReadAll(source, new OfficeLegacyImportLimits { MaxInputBytes = 64 }));
        Assert.Equal("Legacy input exceeds the configured 64 byte limit.", exception.Message);
        Assert.Equal(65, source.BytesRead);
        Assert.True(source.CanRead);
    }

    [Fact]
    public void CancellationRaisedByAReadIsObservedBeforeReturningData() {
        using var cancellation = new CancellationTokenSource();
        using var source = new CountingForwardStream(new byte[64], cancellation);
        Assert.ThrowsAny<OperationCanceledException>(() =>
            OfficeStreamReader.ReadForwardBytes(source, cancellation.Token, 64));
        Assert.Equal(64, source.BytesRead);
        Assert.True(source.CanRead);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task SeekableSnapshotDoesNotRetainASourceReadBuffer(bool asynchronous, bool readRemaining) {
        byte[] bytes = Enumerable.Range(0, 200_017).Select(index => (byte)(index * 13 + index / 251)).ToArray();
        const int startingPosition = 43;
        using var source = new BufferRecordingStream(bytes) { Position = startingPosition };

        byte[] snapshot;
        if (readRemaining) {
            snapshot = asynchronous
                ? await OfficeStreamReader.ReadRemainingBytesAsync(source)
                : OfficeStreamReader.ReadRemainingBytes(source);
        } else {
            snapshot = asynchronous
                ? await OfficeStreamReader.ReadAllBytesAsync(source)
                : OfficeStreamReader.ReadAllBytes(source);
        }

        Assert.Equal(readRemaining ? bytes.Skip(startingPosition).ToArray() : bytes, snapshot);
        Assert.NotEmpty(source.ReadBuffers);
        Assert.DoesNotContain(source.ReadBuffers, buffer => ReferenceEquals(buffer, snapshot));
        Assert.True(source.CanRead);
        Assert.Equal(readRemaining ? bytes.Length : startingPosition, source.Position);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ExpandingSnapshotObservesCancellationAtEndOfInput(bool asynchronous) {
        using var cancellation = new CancellationTokenSource();
        using var source = new ExpandingStream(new byte[211], cancellation);

        if (asynchronous) {
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
                OfficeStreamReader.ReadAllBytesAsync(source, cancellation.Token));
        } else {
            Assert.ThrowsAny<OperationCanceledException>(() =>
                OfficeStreamReader.ReadAllBytes(source, cancellation.Token));
        }

        Assert.True(source.CanRead);
        Assert.Equal(0, source.Position);
    }

    private sealed class ExpandingStream(byte[] bytes, CancellationTokenSource cancellation) : MemoryStream(bytes) {
        public override long Length => 64;

        public override int Read(byte[] buffer, int offset, int count) {
            int read = base.Read(buffer, offset, count);
            if (read == 0) cancellation.Cancel();
            return read;
        }

        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) =>
            Task.FromResult(Read(buffer, offset, count));
    }

    private sealed class BufferRecordingStream(byte[] bytes) : MemoryStream(bytes) {
        internal HashSet<byte[]> ReadBuffers { get; } = new();

        public override int Read(byte[] buffer, int offset, int count) {
            int read = base.Read(buffer, offset, Math.Min(count, 2048));
            if (read > 0) ReadBuffers.Add(buffer);
            return read;
        }

        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) =>
            Task.FromResult(Read(buffer, offset, count));
    }

    private sealed class CountingForwardStream : MemoryStream {
        private readonly CancellationTokenSource? _cancellation;
        internal CountingForwardStream(byte[] bytes, CancellationTokenSource? cancellation = null) : base(bytes) =>
            _cancellation = cancellation;
        internal int BytesRead { get; private set; }
        public override bool CanSeek => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override int Read(byte[] buffer, int offset, int count) {
            int read = base.Read(buffer, offset, count);
            BytesRead += read;
            _cancellation?.Cancel();
            return read;
        }
        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) =>
            Task.FromResult(Read(buffer, offset, count));
    }
}

