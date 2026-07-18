namespace OfficeIMO.Email.Store;

internal sealed class PstBlockReference {
    internal PstBlockReference(ulong bid, long offset, int dataLength, int decodedLength) {
        Bid = bid;
        Offset = offset;
        DataLength = dataLength;
        DecodedLength = decodedLength;
    }

    internal ulong Bid { get; }
    internal long Offset { get; }
    internal int DataLength { get; }
    internal int DecodedLength { get; }
}

internal sealed class PstNodeReference {
    internal PstNodeReference(uint nid, ulong dataBid, ulong subnodeBid, uint parentNid) {
        Nid = nid;
        DataBid = dataBid;
        SubnodeBid = subnodeBid;
        ParentNid = parentNid;
    }

    internal uint Nid { get; }
    internal ulong DataBid { get; }
    internal ulong SubnodeBid { get; }
    internal uint ParentNid { get; }
    internal int Type => (int)(Nid & 0x1F);
}

internal sealed class PstDataTree {
    private readonly IReadOnlyList<byte[]>? _materializedBlocks;
    private readonly Func<IEnumerable<byte[]>>? _blockFactory;
    private readonly int _cacheCapacity;
    private readonly Dictionary<int, LinkedListNode<CachedBlock>> _cache =
        new Dictionary<int, LinkedListNode<CachedBlock>>();
    private readonly LinkedList<CachedBlock> _lru = new LinkedList<CachedBlock>();
    private IEnumerator<byte[]>? _cursor;
    private int _cursorIndex = -1;
    private long? _totalLength;

    internal PstDataTree(IReadOnlyList<byte[]> blocks, long totalLength) {
        _materializedBlocks = blocks ?? throw new ArgumentNullException(nameof(blocks));
        _totalLength = totalLength;
        _cacheCapacity = 0;
    }

    internal PstDataTree(Func<IEnumerable<byte[]>> blockFactory, int cacheCapacity) {
        _blockFactory = blockFactory ?? throw new ArgumentNullException(nameof(blockFactory));
        if (cacheCapacity <= 0) throw new ArgumentOutOfRangeException(nameof(cacheCapacity));
        _cacheCapacity = cacheCapacity;
    }

    internal IReadOnlyList<byte[]> Blocks => _materializedBlocks ?? EnumerateBlocks().ToArray();
    internal long TotalLength {
        get {
            if (_totalLength.HasValue) return _totalLength.Value;
            long total = 0;
            foreach (byte[] block in EnumerateBlocks()) total = checked(total + block.LongLength);
            _totalLength = total;
            return total;
        }
    }

    internal byte[] GetBlock(int index) {
        if (index < 0) throw new ArgumentOutOfRangeException(nameof(index));
        if (_materializedBlocks != null) {
            if (index >= _materializedBlocks.Count) {
                throw new InvalidDataException("A PST data-tree block index is out of range.");
            }
            return _materializedBlocks[index];
        }
        if (_cache.TryGetValue(index, out LinkedListNode<CachedBlock>? cached)) {
            _lru.Remove(cached);
            _lru.AddFirst(cached);
            return cached.Value.Bytes;
        }
        if (index <= _cursorIndex) ResetCursor();
        EnsureCursor();
        while (_cursorIndex < index && _cursor!.MoveNext()) {
            _cursorIndex++;
            AddCached(_cursorIndex, _cursor.Current ??
                throw new InvalidDataException("A PST data-tree block was null."));
        }
        if (_cache.TryGetValue(index, out cached)) return cached.Value.Bytes;
        throw new InvalidDataException("A PST data-tree block index is out of range.");
    }

    internal IEnumerable<byte[]> EnumerateBlocks() =>
        _materializedBlocks ?? _blockFactory!();

    internal byte[] ToArray(long maximum) {
        if (TotalLength > maximum || TotalLength > int.MaxValue) {
            throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxDecodedPropertyBytesPerItem),
                TotalLength, Math.Min(maximum, int.MaxValue));
        }
        var result = new byte[checked((int)TotalLength)];
        int offset = 0;
        foreach (byte[] block in EnumerateBlocks()) {
            Buffer.BlockCopy(block, 0, result, offset, block.Length);
            offset += block.Length;
        }
        return result;
    }

    private void EnsureCursor() {
        if (_cursor == null) _cursor = _blockFactory!().GetEnumerator();
    }

    private void ResetCursor() {
        _cursor?.Dispose();
        _cursor = null;
        _cursorIndex = -1;
    }

    private void AddCached(int index, byte[] bytes) {
        if (_cache.TryGetValue(index, out LinkedListNode<CachedBlock>? existing)) {
            _lru.Remove(existing);
        }
        var entry = new CachedBlock(index, bytes);
        LinkedListNode<CachedBlock> node = _lru.AddFirst(entry);
        _cache[index] = node;
        if (_cache.Count <= _cacheCapacity) return;
        LinkedListNode<CachedBlock>? last = _lru.Last;
        if (last == null) return;
        _lru.RemoveLast();
        _cache.Remove(last.Value.Index);
    }

    private sealed class CachedBlock {
        internal CachedBlock(int index, byte[] bytes) { Index = index; Bytes = bytes; }
        internal int Index { get; }
        internal byte[] Bytes { get; }
    }
}

internal sealed class PstSubnodeReference {
    internal PstSubnodeReference(uint nid, ulong dataBid, ulong subnodeBid) {
        Nid = nid;
        DataBid = dataBid;
        SubnodeBid = subnodeBid;
    }

    internal uint Nid { get; }
    internal ulong DataBid { get; }
    internal ulong SubnodeBid { get; }
    internal int Type => (int)(Nid & 0x1F);
}
