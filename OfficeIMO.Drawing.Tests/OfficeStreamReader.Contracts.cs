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

