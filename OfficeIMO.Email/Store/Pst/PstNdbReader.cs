namespace OfficeIMO.Email.Store;

internal sealed partial class PstNdbReader {
    private readonly Stream _stream;
    private readonly PstHeader _header;
    private readonly EmailStoreReaderOptions _options;
    private readonly CancellationToken _cancellationToken;
    private readonly Dictionary<ulong, PstBlockReference> _blocks = new Dictionary<ulong, PstBlockReference>();
    private readonly Dictionary<uint, PstNodeReference> _nodes = new Dictionary<uint, PstNodeReference>();
    private readonly PstPageCache _pageCache;
    private readonly PstTraversalCounter _traversalCounter;
    private bool _indexesLoaded;

    internal PstNdbReader(Stream stream, PstHeader header, EmailStoreReaderOptions options,
        CancellationToken cancellationToken) {
        _stream = stream;
        _header = header;
        _options = options;
        _cancellationToken = cancellationToken;
        _pageCache = new PstPageCache(options.MaxCachedBTreePages);
        _traversalCounter = new PstTraversalCounter(options.MaxNodeCount);
    }

    internal IReadOnlyDictionary<uint, PstNodeReference> Nodes => _indexesLoaded
        ? _nodes
        : throw new InvalidOperationException("The complete PST indexes have not been loaded.");

    internal bool IsUnicode => _header.IsUnicode;
    internal int HeapAllocationIndexBits => _header.Variant == PstVariant.Unicode4K ? 14 : 11;

    /// <summary>Loads complete BBT and NBT dictionaries for the legacy materializing reader.</summary>
    internal void LoadIndexes() {
        _blocks.Clear();
        _nodes.Clear();
        var budget = CreateBudget();
        TraversePage(_header.BbtRootOffset, 0x80, isBlockTree: true, depth: 0,
            new HashSet<long>(), budget);
        TraversePage(_header.NbtRootOffset, 0x81, isBlockTree: false, depth: 0,
            new HashSet<long>(), budget);
        _indexesLoaded = true;
    }

    /// <summary>Streams NBT leaf entries without retaining the complete node index.</summary>
    internal IEnumerable<PstNodeReference> EnumerateNodes(CancellationToken cancellationToken = default) {
        var budget = CreateBudget(cancellationToken);
        foreach (PstNodeReference node in EnumerateNodePage(_header.NbtRootOffset, depth: 0,
            new HashSet<long>(), budget)) yield return node;
    }

    /// <summary>Resolves one NBT entry by key without loading the complete node index.</summary>
    internal bool TryGetNode(uint nid, out PstNodeReference? node,
        CancellationToken cancellationToken = default) {
        if (_indexesLoaded) return _nodes.TryGetValue(nid, out node);
        return TryFindNode(nid, CreateBudget(cancellationToken), out node);
    }

    internal PstDataTree ReadDataTree(ulong bid, long maximumBytes,
        CancellationToken cancellationToken = default) {
        IReadOnlyList<byte[]> blocks = EnumerateDataBlocks(bid, maximumBytes, cancellationToken).ToArray();
        return new PstDataTree(blocks, blocks.Sum(block => (long)block.Length));
    }

    /// <summary>Opens a bounded, block-on-demand view of a data tree for Heap-on-Node parsing.</summary>
    internal PstDataTree OpenDataTree(ulong bid, long maximumBytes,
        CancellationToken cancellationToken = default) {
        int cacheCapacity = Math.Max(4, Math.Min(64, _options.MaxCachedBTreePages));
        return new PstDataTree(
            () => EnumerateDataBlocks(bid, maximumBytes, cancellationToken), cacheCapacity);
    }

    /// <summary>Streams decoded leaf blocks from a data tree without retaining the complete payload.</summary>
    internal IEnumerable<byte[]> EnumerateDataBlocks(ulong bid, long maximumBytes,
        CancellationToken cancellationToken = default) {
        var byteBudget = new PstDecodedByteBudget(maximumBytes);
        foreach (byte[] block in EnumerateDataBlocksCore(bid, depth: 0, new HashSet<ulong>(),
            CreateBudget(cancellationToken), byteBudget)) yield return block;
    }

    internal IReadOnlyDictionary<uint, PstSubnodeReference> ReadSubnodes(ulong bid,
        CancellationToken cancellationToken = default) {
        var result = new Dictionary<uint, PstSubnodeReference>();
        if (bid == 0) return result;
        ReadSubnodeBlock(bid, result, 0, new HashSet<ulong>(), CreateBudget(cancellationToken));
        return result;
    }

    private void TraversePage(long offset, byte expectedType, bool isBlockTree, int depth,
        HashSet<long> path, PstTraversalBudget budget) {
        budget.CheckDepth(depth);
        if (!path.Add(offset)) throw new InvalidDataException("The PST B-tree contains a page cycle.");
        try {
            PstBTreePage page = ReadPage(offset, expectedType, budget);
            if (page.Level > 0) {
                ValidateIndexEntrySize(page);
                for (int index = 0; index < page.Count; index++) {
                    long childOffset = ReadChildOffset(page, index);
                    TraversePage(childOffset, expectedType, isBlockTree, depth + 1, path, budget);
                }
                return;
            }

            for (int index = 0; index < page.Count; index++) {
                budget.CountStructure();
                int entryOffset = index * page.EntrySize;
                if (isBlockTree) {
                    PstBlockReference block = ReadBlockEntry(page.Bytes, entryOffset, page.EntrySize);
                    ulong key = PstBinary.NormalizeBid(block.Bid);
                    if (!_blocks.ContainsKey(key)) _blocks.Add(key, block);
                } else {
                    PstNodeReference node = ReadNodeEntry(page.Bytes, entryOffset, page.EntrySize);
                    if (!_nodes.ContainsKey(node.Nid)) _nodes.Add(node.Nid, node);
                }
            }
        } finally {
            path.Remove(offset);
        }
    }

    private IEnumerable<PstNodeReference> EnumerateNodePage(long offset, int depth,
        HashSet<long> path, PstTraversalBudget budget) {
        budget.CheckDepth(depth);
        if (!path.Add(offset)) throw new InvalidDataException("The PST NBT contains a page cycle.");
        try {
            PstBTreePage page = ReadPage(offset, 0x81, budget);
            if (page.Level > 0) {
                ValidateIndexEntrySize(page);
                for (int index = 0; index < page.Count; index++) {
                    long childOffset = ReadChildOffset(page, index);
                    foreach (PstNodeReference node in EnumerateNodePage(
                        childOffset, depth + 1, path, budget)) yield return node;
                }
                yield break;
            }

            for (int index = 0; index < page.Count; index++) {
                budget.CountStructure();
                yield return ReadNodeEntry(page.Bytes, index * page.EntrySize, page.EntrySize);
            }
        } finally {
            path.Remove(offset);
        }
    }

    private bool TryFindNode(uint nid, PstTraversalBudget budget, out PstNodeReference? node) {
        long offset = _header.NbtRootOffset;
        var path = new HashSet<long>();
        for (int depth = 0; ; depth++) {
            budget.CheckDepth(depth);
            if (!path.Add(offset)) throw new InvalidDataException("The PST NBT contains a page cycle.");
            PstBTreePage page = ReadPage(offset, 0x81, budget);
            if (page.Level > 0) {
                if (!TrySelectChild(page, nid, isBlockTree: false, out offset)) {
                    node = null;
                    return false;
                }
                continue;
            }

            for (int index = 0; index < page.Count; index++) {
                budget.CountStructure();
                int entryOffset = index * page.EntrySize;
                ulong key = ReadKey(page.Bytes, entryOffset, isBlockTree: false);
                if (key == nid) {
                    node = ReadNodeEntry(page.Bytes, entryOffset, page.EntrySize);
                    return true;
                }
            }
            node = null;
            return false;
        }
    }

    private bool TryFindBlock(ulong normalizedBid, PstTraversalBudget budget,
        out PstBlockReference? block) {
        long offset = _header.BbtRootOffset;
        var path = new HashSet<long>();
        for (int depth = 0; ; depth++) {
            budget.CheckDepth(depth);
            if (!path.Add(offset)) throw new InvalidDataException("The PST BBT contains a page cycle.");
            PstBTreePage page = ReadPage(offset, 0x80, budget);
            if (page.Level > 0) {
                if (!TrySelectChild(page, normalizedBid, isBlockTree: true, out offset)) {
                    block = null;
                    return false;
                }
                continue;
            }

            for (int index = 0; index < page.Count; index++) {
                budget.CountStructure();
                int entryOffset = index * page.EntrySize;
                ulong key = ReadKey(page.Bytes, entryOffset, isBlockTree: true);
                if (key == normalizedBid) {
                    block = ReadBlockEntry(page.Bytes, entryOffset, page.EntrySize);
                    return true;
                }
            }
            block = null;
            return false;
        }
    }

    private bool TrySelectChild(PstBTreePage page, ulong target, bool isBlockTree,
        out long childOffset) {
        ValidateIndexEntrySize(page);
        int selected = -1;
        for (int index = 0; index < page.Count; index++) {
            ulong key = ReadKey(page.Bytes, index * page.EntrySize, isBlockTree);
            if (key > target) break;
            selected = index;
        }
        if (selected < 0) {
            childOffset = 0;
            return false;
        }
        childOffset = ReadChildOffset(page, selected);
        return true;
    }

    private PstBTreePage ReadPage(long offset, byte expectedType, PstTraversalBudget budget) {
        budget.CountStructure();
        byte[] bytes = _pageCache.GetOrAdd(offset,
            () => PstBinary.ReadAt(_stream, offset, _header.PageSize));
        int metadataSize = _header.BTreeMetadataSize;
        int metadataOffset = _header.PageSize - _header.PageTrailerSize - metadataSize;
        int trailerOffset = _header.PageSize - _header.PageTrailerSize;
        if (bytes[trailerOffset] != expectedType || bytes[trailerOffset + 1] != expectedType) {
            throw new InvalidDataException("A PST B-tree page has an unexpected page type.");
        }

        int count;
        int entrySize;
        int level;
        if (_header.Variant == PstVariant.Unicode4K) {
            count = PstBinary.UInt16(bytes, metadataOffset);
            entrySize = bytes[metadataOffset + 4];
            level = bytes[metadataOffset + 5];
        } else {
            count = bytes[metadataOffset];
            entrySize = bytes[metadataOffset + 2];
            level = bytes[metadataOffset + 3];
        }
        if (entrySize <= 0 || checked(count * entrySize) > metadataOffset) {
            throw new InvalidDataException("A PST B-tree page has an invalid entry layout.");
        }
        return new PstBTreePage(bytes, count, entrySize, level);
    }

    private void ValidateIndexEntrySize(PstBTreePage page) {
        int keySize = _header.IsUnicode ? 8 : 4;
        int brefSize = _header.IsUnicode ? 16 : 8;
        if (page.EntrySize < keySize + brefSize) {
            throw new InvalidDataException("A PST B-tree index entry is truncated.");
        }
    }

    private long ReadChildOffset(PstBTreePage page, int index) {
        int keySize = _header.IsUnicode ? 8 : 4;
        int entryOffset = index * page.EntrySize;
        return _header.IsUnicode
            ? checked((long)PstBinary.UInt64(page.Bytes, entryOffset + keySize + 8))
            : PstBinary.UInt32(page.Bytes, entryOffset + keySize + 4);
    }

    private ulong ReadKey(byte[] page, int offset, bool isBlockTree) {
        ulong key = _header.IsUnicode ? PstBinary.UInt64(page, offset) : PstBinary.UInt32(page, offset);
        return isBlockTree ? PstBinary.NormalizeBid(key) : key;
    }

    private PstBlockReference ReadBlockEntry(byte[] page, int offset, int entrySize) {
        int minimum = _header.IsUnicode ? 20 : 12;
        if (entrySize < minimum) throw new InvalidDataException("A PST BBT leaf entry is truncated.");
        ulong bid = _header.IsUnicode ? PstBinary.UInt64(page, offset) : PstBinary.UInt32(page, offset);
        long blockOffset = _header.IsUnicode
            ? checked((long)PstBinary.UInt64(page, offset + 8))
            : PstBinary.UInt32(page, offset + 4);
        int length = PstBinary.UInt16(page, offset + (_header.IsUnicode ? 16 : 8));
        int decodedLength = _header.Variant == PstVariant.Unicode4K
            ? PstBinary.UInt16(page, offset + 18)
            : length;
        if (decodedLength <= 0) decodedLength = length;
        return new PstBlockReference(bid, blockOffset, length, decodedLength);
    }

    private PstNodeReference ReadNodeEntry(byte[] page, int offset, int entrySize) {
        int minimum = _header.IsUnicode ? 28 : 16;
        if (entrySize < minimum) throw new InvalidDataException("A PST NBT leaf entry is truncated.");
        uint nid = _header.IsUnicode ? checked((uint)PstBinary.UInt64(page, offset)) : PstBinary.UInt32(page, offset);
        ulong dataBid = _header.IsUnicode ? PstBinary.UInt64(page, offset + 8) : PstBinary.UInt32(page, offset + 4);
        ulong subBid = _header.IsUnicode ? PstBinary.UInt64(page, offset + 16) : PstBinary.UInt32(page, offset + 8);
        uint parent = PstBinary.UInt32(page, offset + (_header.IsUnicode ? 24 : 12));
        return new PstNodeReference(nid, dataBid, subBid, parent);
    }

    private IEnumerable<byte[]> EnumerateDataBlocksCore(ulong bid, int depth,
        HashSet<ulong> visited, PstTraversalBudget budget, PstDecodedByteBudget byteBudget) {
        budget.CheckDepth(depth);
        ulong key = PstBinary.NormalizeBid(bid);
        if (!visited.Add(key)) throw new InvalidDataException("The PST data tree contains a block cycle.");
        try {
            PstBlockReference block = GetBlock(key, budget);
            byte[] bytes = ReadBlockPayload(block);

            if ((bid & 0x02) == 0) {
                PstCrypt.Decode(bytes, _header.CryptMethod, block.Bid);
                byteBudget.Add(bytes.LongLength);
                yield return bytes;
                yield break;
            }

            if (bytes.Length < 8 || bytes[0] != 0x01 || (bytes[1] != 0x01 && bytes[1] != 0x02)) {
                throw new InvalidDataException("A PST data-tree internal block is malformed.");
            }
            int count = PstBinary.UInt16(bytes, 2);
            int bidSize = _header.BidSize;
            if (8 + checked(count * bidSize) > bytes.Length) {
                throw new InvalidDataException("A PST data-tree internal block is truncated.");
            }
            for (int index = 0; index < count; index++) {
                ulong child = _header.IsUnicode
                    ? PstBinary.UInt64(bytes, 8 + index * bidSize)
                    : PstBinary.UInt32(bytes, 8 + index * bidSize);
                foreach (byte[] childBlock in EnumerateDataBlocksCore(
                    child, depth + 1, visited, budget, byteBudget)) yield return childBlock;
            }
        } finally {
            visited.Remove(key);
        }
    }

    private void ReadSubnodeBlock(ulong bid, Dictionary<uint, PstSubnodeReference> result,
        int depth, HashSet<ulong> visited, PstTraversalBudget budget) {
        budget.CheckDepth(depth);
        ulong key = PstBinary.NormalizeBid(bid);
        if (!visited.Add(key)) throw new InvalidDataException("The PST subnode tree contains a block cycle.");
        try {
            byte[] bytes = ReadBlockPayload(GetBlock(key, budget));
            if (bytes.Length < 4 || bytes[0] != 0x02 || (bytes[1] != 0x00 && bytes[1] != 0x01)) {
                throw new InvalidDataException("A PST subnode block is malformed.");
            }
            int count = PstBinary.UInt16(bytes, 2);
            int headerSize = _header.IsUnicode ? 8 : 4;
            if (bytes[1] == 0) {
                int entrySize = _header.IsUnicode ? 24 : 12;
                if (headerSize + checked(count * entrySize) > bytes.Length) {
                    throw new InvalidDataException("A PST subnode leaf block is truncated.");
                }
                for (int index = 0; index < count; index++) {
                    budget.CountStructure();
                    int offset = headerSize + index * entrySize;
                    uint nid = PstBinary.UInt32(bytes, offset);
                    ulong dataBid = _header.IsUnicode
                        ? PstBinary.UInt64(bytes, offset + 8)
                        : PstBinary.UInt32(bytes, offset + 4);
                    ulong subBid = _header.IsUnicode
                        ? PstBinary.UInt64(bytes, offset + 16)
                        : PstBinary.UInt32(bytes, offset + 8);
                    if (!result.ContainsKey(nid)) result.Add(nid, new PstSubnodeReference(nid, dataBid, subBid));
                }
            } else {
                int entrySize = _header.IsUnicode ? 16 : 8;
                if (headerSize + checked(count * entrySize) > bytes.Length) {
                    throw new InvalidDataException("A PST subnode index block is truncated.");
                }
                for (int index = 0; index < count; index++) {
                    ulong child = _header.IsUnicode
                        ? PstBinary.UInt64(bytes, headerSize + index * entrySize + 8)
                        : PstBinary.UInt32(bytes, headerSize + index * entrySize + 4);
                    ReadSubnodeBlock(child, result, depth + 1, visited, budget);
                }
            }
        } finally {
            visited.Remove(key);
        }
    }

    private byte[] ReadBlockPayload(PstBlockReference block) {
        if (block.DataLength < 0) throw new InvalidDataException("A PST block has an invalid length.");
        if (block.Offset < 0 || block.Offset > _stream.Length - block.DataLength) {
            throw new InvalidDataException("A PST block points outside the source stream.");
        }
        byte[] payload = PstBinary.ReadAt(_stream, block.Offset, block.DataLength);
        return _header.Variant == PstVariant.Unicode4K && block.DecodedLength != payload.Length
            ? PstDeflate.Decode(payload, block.DecodedLength)
            : payload;
    }

    private PstBlockReference GetBlock(ulong normalizedBid, PstTraversalBudget budget) {
        PstBlockReference? block;
        bool found = _indexesLoaded
            ? _blocks.TryGetValue(normalizedBid, out block)
            : TryFindBlock(normalizedBid, budget, out block);
        if (!found || block == null) {
            throw new InvalidDataException(string.Concat("The PST BBT does not contain BID 0x",
                normalizedBid.ToString("X", CultureInfo.InvariantCulture), "."));
        }
        return block;
    }

    private PstTraversalBudget CreateBudget(CancellationToken cancellationToken = default) =>
        new PstTraversalBudget(_options,
            cancellationToken.CanBeCanceled ? cancellationToken : _cancellationToken,
            _traversalCounter);

    private sealed class PstBTreePage {
        internal PstBTreePage(byte[] bytes, int count, int entrySize, int level) {
            Bytes = bytes;
            Count = count;
            EntrySize = entrySize;
            Level = level;
        }

        internal byte[] Bytes { get; }
        internal int Count { get; }
        internal int EntrySize { get; }
        internal int Level { get; }
    }

    private sealed class PstTraversalBudget {
        private readonly EmailStoreReaderOptions _options;
        private readonly CancellationToken _cancellationToken;
        private readonly PstTraversalCounter _counter;

        internal PstTraversalBudget(EmailStoreReaderOptions options,
            CancellationToken cancellationToken,
            PstTraversalCounter counter) {
            _options = options;
            _cancellationToken = cancellationToken;
            _counter = counter;
        }

        internal void CheckDepth(int depth) {
            _cancellationToken.ThrowIfCancellationRequested();
            if (depth > _options.MaxBTreeDepth) {
                throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxBTreeDepth),
                    depth, _options.MaxBTreeDepth);
            }
        }

        internal void CountStructure() {
            _cancellationToken.ThrowIfCancellationRequested();
            _counter.Count();
        }
    }

    private sealed class PstTraversalCounter {
        private readonly int _maximumStructures;
        private int _visitedStructures;

        internal PstTraversalCounter(int maximumStructures) {
            _maximumStructures = maximumStructures;
        }

        internal void Count() {
            int observed = Interlocked.Increment(ref _visitedStructures);
            if (observed > _maximumStructures) {
                throw new EmailStoreLimitExceededException(
                    nameof(EmailStoreReaderOptions.MaxNodeCount),
                    observed, _maximumStructures);
            }
        }
    }

    private sealed class PstDecodedByteBudget {
        private readonly long _maximumBytes;
        private long _decodedBytes;

        internal PstDecodedByteBudget(long maximumBytes) {
            if (maximumBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maximumBytes));
            _maximumBytes = maximumBytes;
        }

        internal void Add(long bytes) {
            _decodedBytes = checked(_decodedBytes + bytes);
            if (_decodedBytes > _maximumBytes) {
                throw new EmailStoreLimitExceededException(
                    nameof(EmailStoreReaderOptions.MaxDecodedPropertyBytesPerItem),
                    _decodedBytes, _maximumBytes);
            }
        }
    }
}
