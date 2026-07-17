namespace OfficeIMO.Email.Store;

internal sealed class PstHeap {
    private readonly PstDataTree _dataTree;
    private readonly IReadOnlyDictionary<uint, PstSubnodeReference> _subnodes;
    private readonly PstNdbReader _ndb;
    private readonly EmailStoreReaderOptions _options;
    private readonly CancellationToken _cancellationToken;
    private readonly byte _clientSignature;
    private readonly uint _userRoot;

    internal PstHeap(PstDataTree dataTree, IReadOnlyDictionary<uint, PstSubnodeReference> subnodes,
        PstNdbReader ndb, EmailStoreReaderOptions options, CancellationToken cancellationToken) {
        _dataTree = dataTree;
        _subnodes = subnodes;
        _ndb = ndb;
        _options = options;
        _cancellationToken = cancellationToken;
        byte[] first = _dataTree.GetBlock(0);
        if (first.Length < 12 || first[2] != 0xEC) {
            throw new InvalidDataException("The PST node does not contain a valid Heap-on-Node header.");
        }
        _clientSignature = first[3];
        _userRoot = PstBinary.UInt32(first, 4);
    }

    internal byte ClientSignature => _clientSignature;
    internal uint UserRoot => _userRoot;

    internal byte[] GetAllocation(uint hid) {
        _cancellationToken.ThrowIfCancellationRequested();
        if (hid == 0) return Array.Empty<byte>();
        if ((hid & 0x1F) != 0) {
            throw new InvalidDataException(string.Concat(
                "Heap allocation identifier 0x", hid.ToString("X8", CultureInfo.InvariantCulture),
                " has non-HID type 0x", (hid & 0x1F).ToString("X2", CultureInfo.InvariantCulture), "."));
        }

        int allocationIndexBits = _ndb.HeapAllocationIndexBits;
        uint allocationIndexMask = (1U << allocationIndexBits) - 1U;
        int allocationIndex = checked((int)((hid >> 5) & allocationIndexMask)) - 1;
        int blockIndex = checked((int)(hid >> (5 + allocationIndexBits)));
        if (allocationIndex < 0) {
            throw new InvalidDataException("A heap allocation identifier is out of range.");
        }

        byte[] block;
        try {
            block = _dataTree.GetBlock(blockIndex);
        } catch (InvalidDataException exception) {
            throw new InvalidDataException(string.Concat(
                "Heap allocation identifier 0x", hid.ToString("X8", CultureInfo.InvariantCulture),
                " references an unavailable HN data block."), exception);
        }
        int mapOffset = PstBinary.UInt16(block, 0);
        PstBinary.Ensure(block, mapOffset, 4);
        int allocationCount = PstBinary.UInt16(block, mapOffset);
        if (allocationIndex >= allocationCount) {
            throw new InvalidDataException(string.Concat(
                "Heap allocation identifier 0x", hid.ToString("X8", CultureInfo.InvariantCulture),
                " requests allocation ", (allocationIndex + 1).ToString(CultureInfo.InvariantCulture),
                " from a page containing ", allocationCount.ToString(CultureInfo.InvariantCulture),
                " allocations (block-bytes=", block.Length.ToString(CultureInfo.InvariantCulture),
                ", map-offset=", mapOffset.ToString(CultureInfo.InvariantCulture),
                ")."));
        }
        int offsetsStart = mapOffset + 4;
        PstBinary.Ensure(block, offsetsStart, checked((allocationCount + 1) * 2));
        int start = PstBinary.UInt16(block, offsetsStart + allocationIndex * 2);
        int end = PstBinary.UInt16(block, offsetsStart + (allocationIndex + 1) * 2);
        if (start < 0 || end < start || end > mapOffset) {
            throw new InvalidDataException("A heap allocation range is invalid.");
        }
        var result = new byte[end - start];
        Buffer.BlockCopy(block, start, result, 0, result.Length);
        return result;
    }

    internal byte[] ResolveHnid(uint hnid, long maximumBytes) {
        if (hnid == 0) return Array.Empty<byte>();
        if ((hnid & 0x1F) == 0) return GetAllocation(hnid);
        if (!_subnodes.TryGetValue(hnid, out PstSubnodeReference? subnode)) {
            throw new InvalidDataException(string.Concat("The PST node does not contain subnode 0x",
                hnid.ToString("X", CultureInfo.InvariantCulture), "."));
        }
        return _ndb.ReadDataTree(subnode.DataBid, maximumBytes, _cancellationToken).ToArray(maximumBytes);
    }

    internal PstDataTree ResolveHnidTree(uint hnid, long maximumBytes) {
        if (hnid == 0) return new PstDataTree(Array.Empty<byte[]>(), 0);
        if ((hnid & 0x1F) == 0) {
            byte[] allocation = GetAllocation(hnid);
            return new PstDataTree(new[] { allocation }, allocation.LongLength);
        }
        if (!_subnodes.TryGetValue(hnid, out PstSubnodeReference? subnode)) {
            throw new InvalidDataException(string.Concat("The PST node does not contain subnode 0x",
                hnid.ToString("X", CultureInfo.InvariantCulture), "."));
        }
        return _ndb.ReadDataTree(subnode.DataBid, maximumBytes, _cancellationToken);
    }

    /// <summary>Streams HNID payload blocks so large table row matrices are not retained as one tree.</summary>
    internal IEnumerable<byte[]> EnumerateHnidBlocks(uint hnid, long maximumBytes) {
        if (hnid == 0) yield break;
        if ((hnid & 0x1F) == 0) {
            yield return GetAllocation(hnid);
            yield break;
        }
        if (!_subnodes.TryGetValue(hnid, out PstSubnodeReference? subnode)) {
            throw new InvalidDataException(string.Concat("The PST node does not contain subnode 0x",
                hnid.ToString("X", CultureInfo.InvariantCulture), "."));
        }
        foreach (byte[] block in _ndb.EnumerateDataBlocks(
            subnode.DataBid, maximumBytes, _cancellationToken)) {
            yield return block;
        }
    }

    internal IEnumerable<byte[]> EnumerateBthLeafRecords(uint rootHid, int keySize, int valueSize,
        int indexLevels) {
        var visited = new HashSet<uint>();
        foreach (byte[] record in EnumerateBthLeafRecordsCore(rootHid, keySize, valueSize,
            indexLevels, 0, visited, parentHid: 0, parentRecordIndex: -1)) yield return record;
    }

    private IEnumerable<byte[]> EnumerateBthLeafRecordsCore(uint hid, int keySize, int valueSize,
        int remainingIndexLevels, int depth, HashSet<uint> visited,
        uint parentHid, int parentRecordIndex) {
        _cancellationToken.ThrowIfCancellationRequested();
        if (depth > _options.MaxBTreeDepth) {
            throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxBTreeDepth),
                depth, _options.MaxBTreeDepth);
        }
        if (!visited.Add(hid)) throw new InvalidDataException("The PST BTree-on-Heap contains a cycle.");
        byte[] allocation;
        try {
            allocation = GetAllocation(hid);
        } catch (InvalidDataException exception) when (parentRecordIndex >= 0) {
            throw new InvalidDataException(string.Concat(
                "BTree-on-Heap allocation 0x", parentHid.ToString("X8", CultureInfo.InvariantCulture),
                " record ", parentRecordIndex.ToString(CultureInfo.InvariantCulture),
                " references invalid child 0x", hid.ToString("X8", CultureInfo.InvariantCulture),
                ": ", exception.Message), exception);
        }
        int recordSize = keySize + (remainingIndexLevels > 0 ? 4 : valueSize);
        if (recordSize <= 0 || allocation.Length % recordSize != 0) {
            throw new InvalidDataException(string.Concat(
                "BTree-on-Heap allocation 0x", hid.ToString("X8", CultureInfo.InvariantCulture),
                " has ", allocation.Length.ToString(CultureInfo.InvariantCulture),
                " bytes, which is not divisible by its ",
                recordSize.ToString(CultureInfo.InvariantCulture), "-byte record size at index level ",
                remainingIndexLevels.ToString(CultureInfo.InvariantCulture), "."));
        }
        int count = allocation.Length / recordSize;
        for (int index = 0; index < count; index++) {
            int offset = index * recordSize;
            if (remainingIndexLevels > 0) {
                uint child = PstBinary.UInt32(allocation, offset + keySize);
                foreach (byte[] record in EnumerateBthLeafRecordsCore(child, keySize, valueSize,
                    remainingIndexLevels - 1, depth + 1, visited, hid, index)) yield return record;
            } else {
                var record = new byte[recordSize];
                Buffer.BlockCopy(allocation, offset, record, 0, recordSize);
                yield return record;
            }
        }
        visited.Remove(hid);
    }
}
