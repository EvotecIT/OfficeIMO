namespace OfficeIMO.Email.Store;

/// <summary>Keeps frequently accessed NDB B-tree pages in a bounded least-recently-used cache.</summary>
internal sealed class PstPageCache {
    private readonly int _capacity;
    private readonly Dictionary<long, CacheEntry> _entries = new Dictionary<long, CacheEntry>();
    private readonly LinkedList<long> _usage = new LinkedList<long>();

    internal PstPageCache(int capacity) {
        if (capacity <= 0) throw new ArgumentOutOfRangeException(nameof(capacity));
        _capacity = capacity;
    }

    internal byte[] GetOrAdd(long offset, Func<byte[]> factory) {
        if (_entries.TryGetValue(offset, out CacheEntry? entry)) {
            _usage.Remove(entry.UsageNode);
            _usage.AddFirst(entry.UsageNode);
            return entry.Bytes;
        }

        byte[] bytes = factory();
        LinkedListNode<long> node = _usage.AddFirst(offset);
        _entries.Add(offset, new CacheEntry(bytes, node));
        if (_entries.Count > _capacity) {
            LinkedListNode<long>? last = _usage.Last;
            if (last != null) {
                _usage.RemoveLast();
                _entries.Remove(last.Value);
            }
        }
        return bytes;
    }

    private sealed class CacheEntry {
        internal CacheEntry(byte[] bytes, LinkedListNode<long> usageNode) {
            Bytes = bytes;
            UsageNode = usageNode;
        }

        internal byte[] Bytes { get; }
        internal LinkedListNode<long> UsageNode { get; }
    }
}
