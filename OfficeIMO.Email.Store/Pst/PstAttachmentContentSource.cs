using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

internal sealed class PstAttachmentContentSource : IEmailContentSource {
    private readonly PstHeap _heap;
    private readonly uint _hnid;
    private readonly long _maximumBytes;
    private readonly EmailStoreSessionLifetime _lifetime;

    internal PstAttachmentContentSource(PstHeap heap, uint hnid, long? length,
        long maximumBytes, EmailStoreSessionLifetime lifetime) {
        _heap = heap ?? throw new ArgumentNullException(nameof(heap));
        _hnid = hnid;
        Length = length;
        _maximumBytes = maximumBytes;
        _lifetime = lifetime ?? throw new ArgumentNullException(nameof(lifetime));
    }

    public long? Length { get; }

    public Stream OpenRead() {
        _lifetime.ThrowIfDisposed();
        return new PstBlockReadStream(
            () => _heap.EnumerateHnidBlocks(_hnid, _maximumBytes), _lifetime);
    }

    public Task<Stream> OpenReadAsync(CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        return Task.FromResult(OpenRead());
    }
}
