namespace OfficeIMO.Email.Store;

internal sealed class PstWriterFile : IDisposable {
    private const int HeaderLength = 564;
    private const int PageSize = 512;
    private const int PageDataLength = 496;
    private const int BlockTrailerLength = 16;
    private const int MaximumBlockPayload = 8176;
    private const long FirstAmapOffset = 0x4400;
    private const long FirstPmapOffset = 0x4600;
    private const long FirstDataOffset = 0x4800;
    private const long AmapInterval = 0x3E000;
    private const long PmapInterval = AmapInterval * 8;

    private readonly FileStream _stream;
    private readonly List<PstWriterBlock> _blocks = new List<PstWriterBlock>();
    private readonly List<PstWriterAllocation> _allocations = new List<PstWriterAllocation>();
    private long _nextOffset = FirstDataOffset;
    private ulong _nextBlockBid = 0x100;
    private ulong _nextPageBid = 0x100;
    private bool _finalized;

    internal PstWriterFile(string path) {
        _stream = new FileStream(path, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.Read,
            128 * 1024, FileOptions.SequentialScan);
        _stream.SetLength(FirstDataOffset);
        RegisterMapPagesThrough(FirstDataOffset);
    }

    internal IReadOnlyList<PstWriterBlock> Blocks => _blocks;

    internal ulong NextBlockBid => _nextBlockBid;

    internal ulong NextPageBid => _nextPageBid;

    internal long Length => _stream.Length;

    internal ulong WriteDataTree(byte[] bytes) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        using (var input = new MemoryStream(bytes, writable: false)) return WriteDataTree(input, bytes.LongLength);
    }

    internal ulong WriteDataTree(Stream input, long length) {
        if (input == null) throw new ArgumentNullException(nameof(input));
        if (!input.CanRead) throw new ArgumentException("The PST data source must be readable.", nameof(input));
        if (length < 0 || length > uint.MaxValue) {
            throw new NotSupportedException("One PST data tree cannot exceed 4 GiB.");
        }
        var leaves = new List<DataTreeBlockReference>();
        long remaining = length;
        if (remaining == 0) return WriteBlock(Array.Empty<byte>(), isInternal: false).Bid;
        var buffer = new byte[MaximumBlockPayload];
        while (remaining > 0) {
            int required = checked((int)Math.Min(buffer.Length, remaining));
            int total = 0;
            while (total < required) {
                int read = input.Read(buffer, total, required - total);
                if (read == 0) throw new EndOfStreamException("The PST data source ended before its declared length.");
                total += read;
            }
            var payload = new byte[total];
            Buffer.BlockCopy(buffer, 0, payload, 0, total);
            leaves.Add(new DataTreeBlockReference(
                WriteBlock(payload, isInternal: false).Bid, checked((uint)payload.Length)));
            remaining -= total;
        }
        return WriteDataTreeIndex(leaves);
    }

    internal ulong WriteDataTree(Stream input, out long length) {
        if (input == null) throw new ArgumentNullException(nameof(input));
        if (!input.CanRead) throw new ArgumentException("The PST data source must be readable.", nameof(input));
        var leaves = new List<DataTreeBlockReference>();
        var buffer = new byte[MaximumBlockPayload];
        long totalLength = 0;
        while (true) {
            int total = 0;
            while (total < buffer.Length) {
                int read = input.Read(buffer, total, buffer.Length - total);
                if (read == 0) break;
                total += read;
            }
            if (total == 0) break;
            totalLength = checked(totalLength + total);
            if (totalLength > uint.MaxValue) {
                throw new NotSupportedException("One PST data tree cannot exceed 4 GiB.");
            }
            var payload = new byte[total];
            Buffer.BlockCopy(buffer, 0, payload, 0, total);
            leaves.Add(new DataTreeBlockReference(
                WriteBlock(payload, isInternal: false).Bid, checked((uint)total)));
            if (total < buffer.Length) break;
        }
        length = totalLength;
        if (leaves.Count == 0) return WriteBlock(Array.Empty<byte>(), isInternal: false).Bid;
        return WriteDataTreeIndex(leaves);
    }

    internal ulong WriteDataTreeBlocks(IReadOnlyList<byte[]> blocks) {
        if (blocks == null) throw new ArgumentNullException(nameof(blocks));
        if (blocks.Count == 0) return WriteBlock(Array.Empty<byte>(), isInternal: false).Bid;
        var leaves = new List<DataTreeBlockReference>(blocks.Count);
        ulong total = 0;
        foreach (byte[] payload in blocks) {
            if (payload == null) throw new ArgumentException("A PST data-tree block cannot be null.", nameof(blocks));
            if (payload.Length > MaximumBlockPayload) {
                throw new ArgumentOutOfRangeException(nameof(blocks), "A PST data-tree block exceeds 8176 bytes.");
            }
            total = checked(total + (uint)payload.Length);
            if (total > uint.MaxValue) throw new NotSupportedException("One PST data tree cannot exceed 4 GiB.");
            leaves.Add(new DataTreeBlockReference(
                WriteBlock(payload, isInternal: false).Bid, checked((uint)payload.Length)));
        }
        return WriteDataTreeIndex(leaves);
    }

    private ulong WriteDataTreeIndex(IReadOnlyList<DataTreeBlockReference> leaves) {
        if (leaves.Count == 1) return leaves[0].Bid;

        const int childrenPerBlock = (MaximumBlockPayload - 8) / 8;
        IReadOnlyList<DataTreeBlockReference> current = leaves;
        int level = 1;
        while (current.Count > 1) {
            if (level > 2) throw new NotSupportedException("The PST data tree exceeds the Unicode XBLOCK depth.");
            var parents = new List<DataTreeBlockReference>((current.Count + childrenPerBlock - 1) / childrenPerBlock);
            for (int offset = 0; offset < current.Count; offset += childrenPerBlock) {
                int count = Math.Min(childrenPerBlock, current.Count - offset);
                var payload = new byte[8 + count * 8];
                payload[0] = 0x01;
                payload[1] = checked((byte)level);
                PstBinary.WriteUInt16(payload, 2, count);
                uint childLength = 0;
                for (int index = 0; index < count; index++) {
                    childLength = checked(childLength + current[offset + index].Length);
                }
                PstBinary.WriteUInt32(payload, 4, childLength);
                for (int index = 0; index < count; index++) {
                    PstBinary.WriteUInt64(payload, 8 + index * 8, current[offset + index].Bid);
                }
                parents.Add(new DataTreeBlockReference(
                    WriteBlock(payload, isInternal: true).Bid, childLength));
            }
            current = parents;
            level++;
        }
        return current[0].Bid;
    }

    internal ulong WriteInternalBlock(byte[] payload) => WriteBlock(payload, isInternal: true).Bid;

    internal PstWriterTreeRoot WriteNodeTree(IReadOnlyList<PstWriterNode> nodes) {
        var ordered = nodes.OrderBy(item => item.Nid).ToArray();
        return WriteBTree(ordered.Length, leafEntrySize: 32, pageType: 0x81,
            (page, pageOffset, itemIndex) => {
                PstWriterNode node = ordered[itemIndex];
                int offset = pageOffset;
                PstBinary.WriteUInt32(page, offset, node.Nid);
                PstBinary.WriteUInt64(page, offset + 8, node.DataBid);
                PstBinary.WriteUInt64(page, offset + 16, node.SubnodeBid);
                PstBinary.WriteUInt32(page, offset + 24, node.ParentNid);
                return node.Nid;
            });
    }

    internal PstWriterTreeRoot WriteBlockTree() {
        PstWriterBlock[] ordered = _blocks.OrderBy(item => PstBinary.NormalizeBid(item.Bid)).ToArray();
        return WriteBTree(ordered.Length, leafEntrySize: 24, pageType: 0x80,
            (page, pageOffset, itemIndex) => {
                PstWriterBlock block = ordered[itemIndex];
                PstBinary.WriteUInt64(page, pageOffset, block.Bid);
                PstBinary.WriteUInt64(page, pageOffset + 8, checked((ulong)block.Offset));
                PstBinary.WriteUInt16(page, pageOffset + 16, block.Length);
                PstBinary.WriteUInt16(page, pageOffset + 18, 1);
                return PstBinary.NormalizeBid(block.Bid);
            });
    }

    internal void FinalizeFile(PstWriterTreeRoot nbt, PstWriterTreeRoot bbt,
        IReadOnlyCollection<PstWriterNode> nodes) {
        if (_finalized) throw new InvalidOperationException("The PST file has already been finalized.");
        long dataPosition = Math.Max(_nextOffset, FirstDataOffset);
        long coverageIndex = (dataPosition - 1 - FirstAmapOffset) / AmapInterval;
        long finalLength = checked(FirstAmapOffset + (coverageIndex + 1) * AmapInterval);
        _stream.SetLength(finalLength);
        RegisterMapPagesThrough(finalLength);
        long freeBytes = WriteAllocationMaps(finalLength);
        WriteDensityListPage();
        WriteHeader(finalLength, freeBytes, nbt, bbt, nodes);
        _stream.Flush(flushToDisk: true);
        _finalized = true;
    }

    public void Dispose() => _stream.Dispose();

    private PstWriterBlock WriteBlock(byte[] payload, bool isInternal) {
        if (payload.Length > MaximumBlockPayload) {
            throw new ArgumentOutOfRangeException(nameof(payload), "A PST block payload exceeds 8176 bytes.");
        }
        int allocationLength = PstBinary.Align(payload.Length + BlockTrailerLength, 64);
        long offset = Allocate(allocationLength, alignment: 64);
        ulong bid = _nextBlockBid | (isInternal ? 0x02UL : 0UL);
        _nextBlockBid = checked(_nextBlockBid + 4);
        var allocation = new byte[allocationLength];
        Buffer.BlockCopy(payload, 0, allocation, 0, payload.Length);
        int trailerOffset = allocationLength - BlockTrailerLength;
        PstBinary.WriteUInt16(allocation, trailerOffset, payload.Length);
        PstBinary.WriteUInt16(allocation, trailerOffset + 2, PstSignature.Compute(offset, bid));
        PstBinary.WriteUInt32(allocation, trailerOffset + 4, PstCrc32.Compute(payload));
        PstBinary.WriteUInt64(allocation, trailerOffset + 8, bid);
        WriteAt(offset, allocation);
        var block = new PstWriterBlock(bid, offset, payload.Length);
        _blocks.Add(block);
        return block;
    }

    private PstWriterTreeRoot WriteBTree(int itemCount, int leafEntrySize, byte pageType,
        Func<byte[], int, int, ulong> writeLeafEntry) {
        int leafCapacity = PageDataLength / leafEntrySize;
        var level = new List<PstWriterPageReference>();
        if (itemCount == 0) {
            level.Add(WriteBTreePage(Array.Empty<PstWriterPageReference>(), 0, leafEntrySize,
                pageType, null, 0, writeLeafEntry));
        } else {
            for (int itemOffset = 0; itemOffset < itemCount; itemOffset += leafCapacity) {
                int count = Math.Min(leafCapacity, itemCount - itemOffset);
                level.Add(WriteBTreePage(Array.Empty<PstWriterPageReference>(), 0, leafEntrySize,
                    pageType, null, count, (page, offset, localIndex) =>
                        writeLeafEntry(page, offset, itemOffset + localIndex)));
            }
        }

        int pageLevel = 1;
        const int indexEntrySize = 24;
        int indexCapacity = PageDataLength / indexEntrySize;
        while (level.Count > 1) {
            var parent = new List<PstWriterPageReference>();
            for (int offset = 0; offset < level.Count; offset += indexCapacity) {
                PstWriterPageReference[] children = level.Skip(offset).Take(indexCapacity).ToArray();
                parent.Add(WriteBTreePage(children, pageLevel, indexEntrySize, pageType,
                    children, children.Length, null));
            }
            level = parent;
            pageLevel++;
        }
        PstWriterPageReference root = level[0];
        return new PstWriterTreeRoot(root.Bid, root.Offset);
    }

    private PstWriterPageReference WriteBTreePage(IReadOnlyList<PstWriterPageReference> keySource,
        int level, int entrySize, byte pageType, IReadOnlyList<PstWriterPageReference>? children,
        int count, Func<byte[], int, int, ulong>? writeLeafEntry) {
        var page = new byte[PageSize];
        ulong firstKey = 0;
        for (int index = 0; index < count; index++) {
            int offset = index * entrySize;
            ulong key;
            if (level == 0) {
                key = writeLeafEntry!(page, offset, index);
            } else {
                PstWriterPageReference child = children![index];
                key = child.Key;
                PstBinary.WriteUInt64(page, offset, key);
                PstBinary.WriteUInt64(page, offset + 8, child.Bid);
                PstBinary.WriteUInt64(page, offset + 16, checked((ulong)child.Offset));
            }
            if (index == 0) firstKey = key;
        }
        page[488] = checked((byte)count);
        page[489] = checked((byte)(PageDataLength / entrySize));
        page[490] = checked((byte)entrySize);
        page[491] = checked((byte)level);
        long pageOffset = Allocate(PageSize, PageSize);
        ulong pageBid = _nextPageBid;
        _nextPageBid = checked(_nextPageBid + 1);
        WritePageTrailer(page, pageType, pageOffset, pageBid, PstSignature.Compute(pageOffset, pageBid));
        WriteAt(pageOffset, page);
        return new PstWriterPageReference(firstKey, pageBid, pageOffset);
    }

    private long Allocate(int length, int alignment) {
        long candidate = (_nextOffset + alignment - 1) & ~(alignment - 1L);
        while (true) {
            long? reserved = FindReservedPageOverlap(candidate, length);
            if (!reserved.HasValue) break;
            candidate = (reserved.Value + PageSize + alignment - 1) & ~(alignment - 1L);
        }
        _nextOffset = checked(candidate + length);
        if (_stream.Length < _nextOffset) _stream.SetLength(_nextOffset);
        _allocations.Add(new PstWriterAllocation(candidate, length));
        RegisterMapPagesThrough(_nextOffset);
        return candidate;
    }

    private static long? FindReservedPageOverlap(long offset, int length) {
        long end = checked(offset + length);
        long firstMapIndex = Math.Max(0, (offset - FirstAmapOffset) / AmapInterval);
        for (long index = firstMapIndex; ; index++) {
            long amap = FirstAmapOffset + index * AmapInterval;
            if (amap >= end) break;
            if (offset < amap + PageSize && end > amap) return amap;
        }
        long firstPmapIndex = Math.Max(0, (offset - FirstPmapOffset) / PmapInterval);
        for (long index = firstPmapIndex; ; index++) {
            long pmap = FirstPmapOffset + index * PmapInterval;
            if (pmap >= end) break;
            if (offset < pmap + PageSize && end > pmap) return pmap;
        }
        return null;
    }

    private void RegisterMapPagesThrough(long end) {
        for (long amap = FirstAmapOffset; amap < end; amap += AmapInterval) {
            RegisterAllocation(amap, PageSize);
        }
        for (long pmap = FirstPmapOffset; pmap < end; pmap += PmapInterval) {
            RegisterAllocation(pmap, PageSize);
        }
    }

    private void RegisterAllocation(long offset, int length) {
        if (_allocations.Any(item => item.Offset == offset && item.Length == length)) return;
        _allocations.Add(new PstWriterAllocation(offset, length));
    }

    private long WriteAllocationMaps(long finalLength) {
        long freeBytes = 0;
        for (long amap = FirstAmapOffset; amap < finalLength; amap += AmapInterval) {
            var page = new byte[PageSize];
            foreach (PstWriterAllocation allocation in _allocations) {
                MarkRange(page, amap, AmapInterval, allocation.Offset, allocation.Length, 64);
            }
            int allocated = CountBits(page, PageDataLength);
            freeBytes = checked(freeBytes + (PageDataLength * 8L - allocated) * 64L);
            WritePageTrailer(page, 0x84, amap, checked((ulong)amap), 0);
            WriteAt(amap, page);
        }
        for (long pmap = FirstPmapOffset; pmap < finalLength; pmap += PmapInterval) {
            var page = new byte[PageSize];
            for (int index = 0; index < PageDataLength; index++) page[index] = 0xFF;
            WritePageTrailer(page, 0x83, pmap, checked((ulong)pmap), 0);
            WriteAt(pmap, page);
        }
        return freeBytes;
    }

    private void WriteDensityListPage() {
        const long densityListOffset = 0x4200;
        var page = new byte[PageSize];
        ulong bid = _nextPageBid;
        _nextPageBid = checked(_nextPageBid + 1);
        WritePageTrailer(page, 0x86, densityListOffset, bid,
            PstSignature.Compute(densityListOffset, bid));
        WriteAt(densityListOffset, page);
    }

    private static void MarkRange(byte[] map, long coverageStart, long coverageLength,
        long allocationOffset, int allocationLength, int unit) {
        long coverageEnd = coverageStart + coverageLength;
        long start = Math.Max(coverageStart, allocationOffset);
        long end = Math.Min(coverageEnd, allocationOffset + allocationLength);
        if (start >= end) return;
        int first = checked((int)((start - coverageStart) / unit));
        int last = checked((int)((end - 1 - coverageStart) / unit));
        for (int index = first; index <= last && index < PageDataLength * 8; index++) {
            map[index / 8] |= checked((byte)(1 << (7 - index % 8)));
        }
    }

    private static int CountBits(byte[] bytes, int count) {
        int result = 0;
        for (int index = 0; index < count; index++) {
            byte value = bytes[index];
            while (value != 0) { result += value & 1; value >>= 1; }
        }
        return result;
    }

    private void WriteHeader(long fileLength, long freeBytes, PstWriterTreeRoot nbt,
        PstWriterTreeRoot bbt, IReadOnlyCollection<PstWriterNode> nodes) {
        var header = new byte[HeaderLength];
        PstBinary.WriteUInt32(header, 0, 0x4E444221);
        PstBinary.WriteUInt16(header, 8, 0x4D53);
        PstBinary.WriteUInt16(header, 10, 23);
        PstBinary.WriteUInt16(header, 12, 19);
        header[14] = 1;
        header[15] = 1;
        PstBinary.WriteUInt64(header, 32, _nextPageBid);
        PstBinary.WriteUInt32(header, 40, 1);
        WriteNidCounters(header, nodes);
        PstBinary.WriteUInt64(header, 184, checked((ulong)fileLength));
        long lastAmap = FirstAmapOffset + ((Math.Max(fileLength - 1, FirstAmapOffset) - FirstAmapOffset) / AmapInterval) * AmapInterval;
        PstBinary.WriteUInt64(header, 192, checked((ulong)lastAmap));
        PstBinary.WriteUInt64(header, 200, checked((ulong)freeBytes));
        PstBinary.WriteUInt64(header, 216, nbt.Bid);
        PstBinary.WriteUInt64(header, 224, checked((ulong)nbt.Offset));
        PstBinary.WriteUInt64(header, 232, bbt.Bid);
        PstBinary.WriteUInt64(header, 240, checked((ulong)bbt.Offset));
        header[248] = 0x02;
        for (int index = 256; index < 512; index++) header[index] = 0xFF;
        header[512] = 0x80;
        header[513] = 0x00;
        PstBinary.WriteUInt64(header, 516, _nextBlockBid);
        PstBinary.WriteUInt32(header, 524, PstCrc32.Compute(header, 8, 516));
        PstBinary.WriteUInt32(header, 4, PstCrc32.Compute(header, 8, 471));
        WriteAt(0, header);
    }

    private static void WriteNidCounters(byte[] header, IReadOnlyCollection<PstWriterNode> nodes) {
        var starts = Enumerable.Repeat(0x400U, 32).ToArray();
        starts[2] = 0x400;
        starts[3] = 0x4000;
        starts[4] = 0x10000;
        starts[8] = 0x8000;
        foreach (PstWriterNode node in nodes) {
            int type = checked((int)(node.Nid & 0x1F));
            uint index = node.Nid >> 5;
            if (index > starts[type]) starts[type] = index;
        }
        for (int index = 0; index < starts.Length; index++) {
            PstBinary.WriteUInt32(header, 44 + index * 4, starts[index]);
        }
    }

    private static void WritePageTrailer(byte[] page, byte pageType, long offset, ulong bid, ushort signature) {
        page[496] = pageType;
        page[497] = pageType;
        PstBinary.WriteUInt16(page, 498, signature);
        PstBinary.WriteUInt32(page, 500, PstCrc32.Compute(page, 0, PageDataLength));
        PstBinary.WriteUInt64(page, 504, bid);
    }

    private void WriteAt(long offset, byte[] bytes) {
        _stream.Position = offset;
        _stream.Write(bytes, 0, bytes.Length);
    }

    private readonly struct DataTreeBlockReference {
        internal DataTreeBlockReference(ulong bid, uint length) {
            Bid = bid;
            Length = length;
        }

        internal ulong Bid { get; }
        internal uint Length { get; }
    }
}
