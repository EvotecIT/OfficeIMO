namespace OfficeIMO.Studio.Features.Reader;

/// <summary>
/// Serializes access to the current PDF renderer and retains a bounded LRU cache of encoded pages.
/// </summary>
internal sealed class PageRenderCoordinator : IDisposable {
    private readonly Func<int, double, CancellationToken, Task<PdfRenderedPage>> _renderPage;
    private readonly int _maximumEntries;
    private readonly long _maximumBytes;
    private readonly SemaphoreSlim _renderGate = new(1, 1);
    private readonly CancellationTokenSource _shutdown = new();
    private readonly object _sync = new();
    private readonly Dictionary<RenderKey, CacheEntry> _cache = new();
    private long _cacheBytes;
    private long _accessSequence;
    private volatile bool _disposed;

    internal PageRenderCoordinator(
        Func<int, double, CancellationToken, Task<PdfRenderedPage>> renderPage,
        int maximumEntries = 16,
        long maximumBytes = 96L * 1024L * 1024L) {
        _renderPage = renderPage ?? throw new ArgumentNullException(nameof(renderPage));
        if (maximumEntries <= 0) throw new ArgumentOutOfRangeException(nameof(maximumEntries));
        if (maximumBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maximumBytes));
        _maximumEntries = maximumEntries;
        _maximumBytes = maximumBytes;
    }

    internal int CachedEntryCount {
        get {
            lock (_sync) return _cache.Count;
        }
    }

    internal long CachedByteCount {
        get {
            lock (_sync) return _cacheBytes;
        }
    }

    internal async Task<PdfRenderedPage> GetPageAsync(
        int pageNumber,
        double scale,
        CancellationToken cancellationToken) {
        ThrowIfDisposed();
        var key = RenderKey.Create(pageNumber, scale);
        if (TryGetCached(key, out PdfRenderedPage? cached)) {
            return cached!;
        }

        using var linkedCancellation = CancellationTokenSource.CreateLinkedTokenSource(
            cancellationToken,
            _shutdown.Token);
        CancellationToken token = linkedCancellation.Token;

        await _renderGate.WaitAsync(token).ConfigureAwait(false);
        try {
            ThrowIfDisposed();
            if (TryGetCached(key, out cached)) {
                return cached!;
            }

            PdfRenderedPage rendered = await _renderPage(pageNumber, key.Scale, token).ConfigureAwait(false);
            token.ThrowIfCancellationRequested();
            AddToCache(key, rendered);
            return rendered;
        } finally {
            _renderGate.Release();
        }
    }

    internal void Clear() {
        lock (_sync) {
            _cache.Clear();
            _cacheBytes = 0;
        }
    }

    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        _shutdown.Cancel();
        Clear();
        _shutdown.Dispose();
        // An in-flight renderer still releases this gate from its finally block after observing
        // shutdown cancellation. Disposing it here would replace cancellation with an
        // ObjectDisposedException on that valid close path. SemaphoreSlim owns no wait handle
        // unless one is explicitly requested, which this coordinator never does.
    }

    private bool TryGetCached(RenderKey key, out PdfRenderedPage? page) {
        lock (_sync) {
            if (_cache.TryGetValue(key, out CacheEntry? entry)) {
                entry.LastAccess = ++_accessSequence;
                page = entry.Page;
                return true;
            }
        }

        page = null;
        return false;
    }

    private void AddToCache(RenderKey key, PdfRenderedPage page) {
        if (page.ByteLength > _maximumBytes) return;

        lock (_sync) {
            if (_disposed) return;

            if (_cache.TryGetValue(key, out CacheEntry? existing)) {
                _cacheBytes -= existing.Page.ByteLength;
            }

            _cache[key] = new CacheEntry(page, ++_accessSequence);
            _cacheBytes += page.ByteLength;

            while (_cache.Count > _maximumEntries || _cacheBytes > _maximumBytes) {
                KeyValuePair<RenderKey, CacheEntry> oldest = _cache.Aggregate(
                    static (left, right) => left.Value.LastAccess <= right.Value.LastAccess ? left : right);
                _cache.Remove(oldest.Key);
                _cacheBytes -= oldest.Value.Page.ByteLength;
            }
        }
    }

    private void ThrowIfDisposed() {
        ObjectDisposedException.ThrowIf(_disposed, this);
    }

    private readonly record struct RenderKey(int PageNumber, int ScalePercent) {
        internal double Scale => ScalePercent / 100D;

        internal static RenderKey Create(int pageNumber, double scale) {
            if (pageNumber <= 0) throw new ArgumentOutOfRangeException(nameof(pageNumber));
            if (double.IsNaN(scale) || double.IsInfinity(scale) || scale <= 0) {
                throw new ArgumentOutOfRangeException(nameof(scale));
            }

            double bounded = Math.Clamp(scale, 0.5D, 2.5D);
            int percent = checked((int)Math.Round(bounded * 4D, MidpointRounding.AwayFromZero) * 25);
            return new RenderKey(pageNumber, percent);
        }
    }

    private sealed class CacheEntry {
        internal CacheEntry(PdfRenderedPage page, long lastAccess) {
            Page = page;
            LastAccess = lastAccess;
        }

        internal PdfRenderedPage Page { get; }

        internal long LastAccess { get; set; }
    }
}
