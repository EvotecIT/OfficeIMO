namespace OfficeIMO.Drawing;

/// <summary>
/// Lets process-wide caches give memory back when the runtime reports high memory load. A
/// finalizable sentinel re-registers itself after each collection and acts only once per full
/// (generation 2) collection, so there is no timer or background thread. Trimming only drops
/// cache entries; work that already holds an entry keeps using it.
/// </summary>
internal static class OfficeMemoryPressure {
    private static readonly object Sync = new();
    private static readonly List<WeakReference<IOfficeMemoryTrimmable>> Trimmables = new();
    private static int _sentinelStarted;
    private static int _lastGen2Count = GC.CollectionCount(2);

    /// <summary>Registers a cache for trimming without keeping it alive.</summary>
    internal static void Register(IOfficeMemoryTrimmable trimmable) {
        lock (Sync) Trimmables.Add(new WeakReference<IOfficeMemoryTrimmable>(trimmable));
        if (Interlocked.Exchange(ref _sentinelStarted, 1) == 0) _ = new Gen2Sentinel();
    }

    /// <summary>Halves every registered cache, oldest entries first.</summary>
    internal static void TrimAll() {
        List<IOfficeMemoryTrimmable> live;
        lock (Sync) {
            live = new List<IOfficeMemoryTrimmable>(Trimmables.Count);
            Trimmables.RemoveAll(reference => {
                if (!reference.TryGetTarget(out IOfficeMemoryTrimmable? target)) return true;
                live.Add(target);
                return false;
            });
        }
        foreach (IOfficeMemoryTrimmable trimmable in live) trimmable.TrimHalf();
    }

    private static void OnCollection() {
        int gen2 = GC.CollectionCount(2);
        if (Interlocked.Exchange(ref _lastGen2Count, gen2) == gen2) return;
        if (IsMemoryLoadHigh()) TrimAll();
    }

#if NETCOREAPP3_0_OR_GREATER
    private static bool IsMemoryLoadHigh() {
        GCMemoryInfo info = GC.GetGCMemoryInfo();
        return info.HighMemoryLoadThresholdBytes > 0 && info.MemoryLoadBytes >= info.HighMemoryLoadThresholdBytes;
    }
#else
    // The netstandard2.0 build reads the same counters when its runtime provides them. A runtime
    // without memory-load information cannot establish high memory load and leaves the caches alone.
    private static readonly Func<(long Load, long Threshold)>? ReadMemoryLoad = CreateMemoryLoadReader();

    private static bool IsMemoryLoadHigh() {
        if (ReadMemoryLoad == null) return false;
        (long load, long threshold) = ReadMemoryLoad();
        return threshold > 0 && load >= threshold;
    }

    private static Func<(long Load, long Threshold)>? CreateMemoryLoadReader() {
        try {
            System.Reflection.MethodInfo? method = typeof(GC).GetMethod("GetGCMemoryInfo", Type.EmptyTypes);
            Type? infoType = method?.ReturnType;
            System.Reflection.PropertyInfo? load = infoType?.GetProperty("MemoryLoadBytes");
            System.Reflection.PropertyInfo? threshold = infoType?.GetProperty("HighMemoryLoadThresholdBytes");
            if (method == null || load == null || threshold == null) return null;
            return () => {
                object info = method.Invoke(null, null)!;
                return ((long)load.GetValue(info)!, (long)threshold.GetValue(info)!);
            };
        } catch (Exception exception) when (!(exception is OutOfMemoryException)) {
            return null;
        }
    }
#endif

    private sealed class Gen2Sentinel {
        ~Gen2Sentinel() {
            if (Environment.HasShutdownStarted) return;
            try {
                OnCollection();
            } catch (Exception) {
                // Trimming is best effort and must not terminate the finalizer thread.
            } finally {
                GC.ReRegisterForFinalize(this);
            }
        }
    }
}

/// <summary>A cache that can drop its older half when memory is scarce.</summary>
internal interface IOfficeMemoryTrimmable {
    void TrimHalf();
}

/// <summary>
/// Thread-safe least-recently-used cache bounded by the summed size of its values. It registers
/// with <see cref="OfficeMemoryPressure"/> and loses its older half under high memory load.
/// </summary>
internal sealed class OfficeSizedLruCache<TKey, TValue> : IOfficeMemoryTrimmable where TKey : notnull {
    private readonly long _maximumSize;
    private readonly Func<TKey, TValue, long> _size;
    private readonly Dictionary<TKey, LinkedListNode<Entry>> _entries;
    private readonly LinkedList<Entry> _order = new();
    private long _retainedSize;

    internal OfficeSizedLruCache(long maximumSize, Func<TKey, TValue, long> size, IEqualityComparer<TKey>? comparer = null) {
        _maximumSize = maximumSize;
        _size = size;
        _entries = new Dictionary<TKey, LinkedListNode<Entry>>(comparer ?? EqualityComparer<TKey>.Default);
        OfficeMemoryPressure.Register(this);
    }

    internal int Count {
        get { lock (_entries) return _entries.Count; }
    }

    internal long RetainedSize {
        get { lock (_entries) return _retainedSize; }
    }

    internal bool TryGet(TKey key, out TValue value) {
        lock (_entries) {
            if (_entries.TryGetValue(key, out LinkedListNode<Entry>? node)) {
                _order.Remove(node);
                _order.AddFirst(node);
                value = node.Value.Value;
                return true;
            }
        }
        value = default!;
        return false;
    }

    /// <summary>Adds a value unless the key is already present, evicting the oldest entries to fit.</summary>
    internal void Add(TKey key, TValue value) {
        long size = _size(key, value);
        if (size > _maximumSize) return;
        lock (_entries) {
            if (_entries.ContainsKey(key)) return;
            while (_retainedSize + size > _maximumSize && _order.Last != null) RemoveOldest();
            _entries.Add(key, _order.AddFirst(new Entry(key, value, size)));
            _retainedSize += size;
        }
    }

    public void TrimHalf() {
        lock (_entries) {
            for (int remove = (_entries.Count + 1) / 2; remove > 0; remove--) RemoveOldest();
        }
    }

    private void RemoveOldest() {
        LinkedListNode<Entry> oldest = _order.Last!;
        _order.RemoveLast();
        _entries.Remove(oldest.Value.Key);
        _retainedSize -= oldest.Value.Size;
    }

    private sealed class Entry {
        internal Entry(TKey key, TValue value, long size) {
            Key = key;
            Value = value;
            Size = size;
        }

        internal TKey Key { get; }
        internal TValue Value { get; }
        internal long Size { get; }
    }
}
