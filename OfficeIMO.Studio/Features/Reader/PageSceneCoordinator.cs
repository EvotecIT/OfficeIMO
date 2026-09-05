namespace OfficeIMO.Studio.Features.Reader;

/// <summary>
/// Serializes OfficeIMO page-scene construction and retains a bounded least-recently-used scene cache.
/// </summary>
internal sealed class PageSceneCoordinator : IDisposable {
    private readonly Func<int, CancellationToken, Task<PdfPageScene>> _loadPage;
    private readonly int _maximumEntries;
    private readonly int _maximumElements;
    private readonly long _maximumBytes;
    private readonly SemaphoreSlim _loadGate = new(1, 1);
    private readonly CancellationTokenSource _shutdown = new();
    private readonly object _sync = new();
    private readonly Dictionary<int, CacheEntry> _cache = new();
    private int _cachedElements;
    private long _cachedBytes;
    private long _accessSequence;
    private volatile bool _disposed;

    internal PageSceneCoordinator(
        Func<int, CancellationToken, Task<PdfPageScene>> loadPage,
        int maximumEntries = 12,
        int maximumElements = 250_000,
        long maximumBytes = 128L * 1024L * 1024L) {
        _loadPage = loadPage ?? throw new ArgumentNullException(nameof(loadPage));
        if (maximumEntries <= 0) throw new ArgumentOutOfRangeException(nameof(maximumEntries));
        if (maximumElements <= 0) throw new ArgumentOutOfRangeException(nameof(maximumElements));
        if (maximumBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maximumBytes));
        _maximumEntries = maximumEntries;
        _maximumElements = maximumElements;
        _maximumBytes = maximumBytes;
    }

    internal int CachedEntryCount {
        get {
            lock (_sync) return _cache.Count;
        }
    }

    internal int CachedElementCount {
        get {
            lock (_sync) return _cachedElements;
        }
    }


    internal long CachedByteCount {
        get {
            lock (_sync) return _cachedBytes;
        }
    }

    internal async Task<PdfPageScene> GetPageAsync(int pageNumber, CancellationToken cancellationToken) {
        ThrowIfDisposed();
        if (pageNumber <= 0) throw new ArgumentOutOfRangeException(nameof(pageNumber));
        if (TryGetCached(pageNumber, out PdfPageScene? cached)) return cached!;

        using var linkedCancellation = CancellationTokenSource.CreateLinkedTokenSource(
            cancellationToken,
            _shutdown.Token);
        CancellationToken token = linkedCancellation.Token;

        await _loadGate.WaitAsync(token).ConfigureAwait(false);
        try {
            ThrowIfDisposed();
            if (TryGetCached(pageNumber, out cached)) return cached!;

            PdfPageScene scene = await _loadPage(pageNumber, token).ConfigureAwait(false);
            token.ThrowIfCancellationRequested();
            AddToCache(pageNumber, scene);
            return scene;
        } finally {
            _loadGate.Release();
        }
    }

    internal void Clear() {
        lock (_sync) {
            _cache.Clear();
            _cachedElements = 0;
            _cachedBytes = 0;
        }
    }

    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        _shutdown.Cancel();
        Clear();
        _shutdown.Dispose();
    }

    private bool TryGetCached(int pageNumber, out PdfPageScene? scene) {
        lock (_sync) {
            if (_cache.TryGetValue(pageNumber, out CacheEntry? entry)) {
                entry.LastAccess = ++_accessSequence;
                scene = entry.Scene;
                return true;
            }
        }

        scene = null;
        return false;
    }

    private void AddToCache(int pageNumber, PdfPageScene scene) {
        if (scene.ElementCount > _maximumElements || scene.EstimatedBytes > _maximumBytes) return;

        lock (_sync) {
            if (_disposed) return;
            if (_cache.TryGetValue(pageNumber, out CacheEntry? existing)) {
                _cachedElements -= existing.Scene.ElementCount;
                _cachedBytes -= existing.Scene.EstimatedBytes;
            }

            _cache[pageNumber] = new CacheEntry(scene, ++_accessSequence);
            _cachedElements += scene.ElementCount;
            _cachedBytes += scene.EstimatedBytes;

            while (_cache.Count > _maximumEntries || _cachedElements > _maximumElements || _cachedBytes > _maximumBytes) {
                KeyValuePair<int, CacheEntry> oldest = _cache.Aggregate(
                    static (left, right) => left.Value.LastAccess <= right.Value.LastAccess ? left : right);
                _cache.Remove(oldest.Key);
                _cachedElements -= oldest.Value.Scene.ElementCount;
                _cachedBytes -= oldest.Value.Scene.EstimatedBytes;
            }
        }
    }

    private void ThrowIfDisposed() => ObjectDisposedException.ThrowIf(_disposed, this);

    private sealed class CacheEntry {
        internal CacheEntry(PdfPageScene scene, long lastAccess) {
            Scene = scene;
            LastAccess = lastAccess;
        }

        internal PdfPageScene Scene { get; }
        internal long LastAccess { get; set; }
    }
}
