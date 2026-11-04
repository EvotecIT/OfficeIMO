using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>
/// A thread-safe memo with an approximate entry limit. <see cref="ConcurrentDictionary{TKey,TValue}.Count"/>
/// takes every bucket lock, so the limit is tracked with its own counter instead.
/// </summary>
internal sealed class OfficeBoundedCache<TKey, TValue> where TKey : notnull {
    private readonly ConcurrentDictionary<TKey, TValue> _items;
    private readonly int _limit;
    private int _count;

    internal OfficeBoundedCache(int limit, IEqualityComparer<TKey>? comparer = null) {
        _limit = limit;
        _items = comparer == null ? new ConcurrentDictionary<TKey, TValue>() : new ConcurrentDictionary<TKey, TValue>(comparer);
    }

    internal bool IsFull => Volatile.Read(ref _count) >= _limit;

    internal bool TryGetValue(TKey key, out TValue value) => _items.TryGetValue(key, out value!);

    internal void TryAdd(TKey key, TValue value) {
        if (IsFull) return;
        if (_items.TryAdd(key, value)) Interlocked.Increment(ref _count);
    }
}
