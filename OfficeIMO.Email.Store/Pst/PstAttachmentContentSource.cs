using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

internal sealed class PstAttachmentContentSource : IEmailContentSource {
    private readonly PstHeap _heap;
    private readonly uint _hnid;
    private readonly long _maximumBytes;
    private readonly long _reservedBytes;
    private readonly PstAttachmentAggregateBudget _aggregateBudget;
    private readonly EmailStoreSessionLifetime _lifetime;
    private readonly object _gate = new object();
    private long _maximumObservedBytes;

    internal PstAttachmentContentSource(PstHeap heap, uint hnid, long? length,
        long maximumBytes, PstAttachmentAggregateBudget aggregateBudget,
        EmailStoreSessionLifetime lifetime) {
        _heap = heap ?? throw new ArgumentNullException(nameof(heap));
        _hnid = hnid;
        Length = length;
        _maximumBytes = maximumBytes;
        _reservedBytes = length.GetValueOrDefault();
        _aggregateBudget = aggregateBudget ?? throw new ArgumentNullException(nameof(aggregateBudget));
        _lifetime = lifetime ?? throw new ArgumentNullException(nameof(lifetime));
    }

    public long? Length { get; }

    public Stream OpenRead() {
        _lifetime.ThrowIfDisposed();
        long streamBytes = 0;
        return new PstBlockReadStream(
            () => _heap.EnumerateHnidBlocks(_hnid, _maximumBytes), _lifetime,
            bytesRead => {
                streamBytes = checked(streamBytes + bytesRead);
                ObserveBytes(streamBytes);
            });
    }

    public Task<Stream> OpenReadAsync(CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        return Task.FromResult(OpenRead());
    }

    private void ObserveBytes(long observedBytes) {
        lock (_gate) {
            if (observedBytes > _maximumBytes) {
                throw new EmailStoreLimitExceededException(
                    nameof(EmailStoreReaderOptions.MaxAttachmentBytes),
                    observedBytes, _maximumBytes);
            }
            if (observedBytes <= _maximumObservedBytes) return;
            long previousCharged = Math.Max(_reservedBytes, _maximumObservedBytes);
            long nextCharged = Math.Max(_reservedBytes, observedBytes);
            _aggregateBudget.Add(nextCharged - previousCharged);
            _maximumObservedBytes = observedBytes;
        }
    }
}

internal sealed class PstAttachmentAggregateBudget {
    private readonly long _maximumBytes;
    private readonly object _gate = new object();
    private long _totalBytes;

    internal PstAttachmentAggregateBudget(long maximumBytes) {
        if (maximumBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maximumBytes));
        _maximumBytes = maximumBytes;
    }

    internal void Add(long bytes) {
        if (bytes < 0) throw new ArgumentOutOfRangeException(nameof(bytes));
        if (bytes == 0) return;
        lock (_gate) {
            if (bytes > _maximumBytes - _totalBytes) {
                long actual = bytes > long.MaxValue - _totalBytes
                    ? long.MaxValue
                    : _totalBytes + bytes;
                throw new EmailStoreLimitExceededException(
                    nameof(EmailStoreReaderOptions.MaxTotalAttachmentBytes),
                    actual, _maximumBytes);
            }
            _totalBytes += bytes;
        }
    }
}
