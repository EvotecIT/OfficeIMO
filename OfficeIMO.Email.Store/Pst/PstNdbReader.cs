namespace OfficeIMO.Email.Store;

internal sealed class PstNdbReader {
    private readonly Stream _stream;
    private readonly PstHeader _header;
    private readonly EmailStoreReaderOptions _options;
    private readonly CancellationToken _cancellationToken;
    private readonly Dictionary<ulong, PstBlockReference> _blocks = new Dictionary<ulong, PstBlockReference>();
    private readonly Dictionary<uint, PstNodeReference> _nodes = new Dictionary<uint, PstNodeReference>();
    private int _visitedStructures;

    internal PstNdbReader(Stream stream, PstHeader header, EmailStoreReaderOptions options,
        CancellationToken cancellationToken) {
        _stream = stream;
        _header = header;
        _options = options;
        _cancellationToken = cancellationToken;
    }

    internal IReadOnlyDictionary<uint, PstNodeReference> Nodes => _nodes;
    internal bool IsUnicode => _header.IsUnicode;

    internal void LoadIndexes() {
        TraversePage(_header.BbtRootOffset, 0x80, isBlockTree: true, depth: 0, new HashSet<long>());
        TraversePage(_header.NbtRootOffset, 0x81, isBlockTree: false, depth: 0, new HashSet<long>());
    }

    internal PstDataTree ReadDataTree(ulong bid, long maximumBytes) {
        var blocks = new List<byte[]>();
        long total = 0;
        ReadDataTreeCore(bid, blocks, ref total, maximumBytes, 0, new HashSet<ulong>());
        return new PstDataTree(blocks, total);
    }

    internal IReadOnlyDictionary<uint, PstSubnodeReference> ReadSubnodes(ulong bid) {
        var result = new Dictionary<uint, PstSubnodeReference>();
        if (bid == 0) return result;
        ReadSubnodeBlock(bid, result, 0, new HashSet<ulong>());
        return result;
    }

    private void TraversePage(long offset, byte expectedType, bool isBlockTree, int depth, HashSet<long> visited) {
        CountStructure();
        if (depth > _options.MaxBTreeDepth) {
            throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxBTreeDepth),
                depth, _options.MaxBTreeDepth);
        }
        if (!visited.Add(offset)) throw new InvalidDataException("The PST B-tree contains a page cycle.");

        byte[] page = PstBinary.ReadAt(_stream, offset, _header.PageSize);
        int metadataSize = _header.BTreeMetadataSize;
        int metadataOffset = _header.PageSize - _header.PageTrailerSize - metadataSize;
        int trailerOffset = _header.PageSize - _header.PageTrailerSize;
        if (page[trailerOffset] != expectedType || page[trailerOffset + 1] != expectedType) {
            throw new InvalidDataException("A PST B-tree page has an unexpected page type.");
        }

        int count;
        int entrySize;
        int level;
        if (_header.Variant == PstVariant.Unicode4K) {
            count = PstBinary.UInt16(page, metadataOffset);
            entrySize = page[metadataOffset + 4];
            level = page[metadataOffset + 5];
        } else {
            count = page[metadataOffset];
            entrySize = page[metadataOffset + 2];
            level = page[metadataOffset + 3];
        }
        if (entrySize <= 0 || checked(count * entrySize) > metadataOffset) {
            throw new InvalidDataException("A PST B-tree page has an invalid entry layout.");
        }

        if (level > 0) {
            int keySize = _header.IsUnicode ? 8 : 4;
            int brefSize = _header.IsUnicode ? 16 : 8;
            if (entrySize < keySize + brefSize) throw new InvalidDataException("A PST B-tree index entry is truncated.");
            for (int index = 0; index < count; index++) {
                int entryOffset = index * entrySize;
                long childOffset = _header.IsUnicode
                    ? checked((long)PstBinary.UInt64(page, entryOffset + keySize + 8))
                    : PstBinary.UInt32(page, entryOffset + keySize + 4);
                TraversePage(childOffset, expectedType, isBlockTree, depth + 1, visited);
            }
            return;
        }

        for (int index = 0; index < count; index++) {
            CountStructure();
            int entryOffset = index * entrySize;
            if (isBlockTree) ReadBlockEntry(page, entryOffset, entrySize);
            else ReadNodeEntry(page, entryOffset, entrySize);
        }
    }

    private void ReadBlockEntry(byte[] page, int offset, int entrySize) {
        int minimum = _header.IsUnicode ? 20 : 12;
        if (entrySize < minimum) throw new InvalidDataException("A PST BBT leaf entry is truncated.");
        ulong bid = _header.IsUnicode ? PstBinary.UInt64(page, offset) : PstBinary.UInt32(page, offset);
        long blockOffset = _header.IsUnicode
            ? checked((long)PstBinary.UInt64(page, offset + 8))
            : PstBinary.UInt32(page, offset + 4);
        int length = PstBinary.UInt16(page, offset + (_header.IsUnicode ? 16 : 8));
        ulong key = PstBinary.NormalizeBid(bid);
        if (!_blocks.ContainsKey(key)) _blocks.Add(key, new PstBlockReference(bid, blockOffset, length));
    }

    private void ReadNodeEntry(byte[] page, int offset, int entrySize) {
        int minimum = _header.IsUnicode ? 28 : 16;
        if (entrySize < minimum) throw new InvalidDataException("A PST NBT leaf entry is truncated.");
        uint nid = _header.IsUnicode ? checked((uint)PstBinary.UInt64(page, offset)) : PstBinary.UInt32(page, offset);
        ulong dataBid = _header.IsUnicode ? PstBinary.UInt64(page, offset + 8) : PstBinary.UInt32(page, offset + 4);
        ulong subBid = _header.IsUnicode ? PstBinary.UInt64(page, offset + 16) : PstBinary.UInt32(page, offset + 8);
        uint parent = PstBinary.UInt32(page, offset + (_header.IsUnicode ? 24 : 12));
        if (!_nodes.ContainsKey(nid)) _nodes.Add(nid, new PstNodeReference(nid, dataBid, subBid, parent));
    }

    private void ReadDataTreeCore(ulong bid, List<byte[]> result, ref long total, long maximumBytes,
        int depth, HashSet<ulong> visited) {
        CountStructure();
        if (depth > _options.MaxBTreeDepth) {
            throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxBTreeDepth),
                depth, _options.MaxBTreeDepth);
        }
        ulong key = PstBinary.NormalizeBid(bid);
        if (!visited.Add(key)) throw new InvalidDataException("The PST data tree contains a block cycle.");
        PstBlockReference block = GetBlock(key);
        byte[] bytes = ReadBlockPayload(block);

        if ((bid & 0x02) == 0) {
            PstCrypt.Decode(bytes, _header.CryptMethod, block.Bid);
            total = checked(total + bytes.Length);
            if (total > maximumBytes) {
                throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxDecodedPropertyBytesPerItem),
                    total, maximumBytes);
            }
            result.Add(bytes);
            visited.Remove(key);
            return;
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
            ReadDataTreeCore(child, result, ref total, maximumBytes, depth + 1, visited);
        }
        visited.Remove(key);
    }

    private void ReadSubnodeBlock(ulong bid, Dictionary<uint, PstSubnodeReference> result,
        int depth, HashSet<ulong> visited) {
        CountStructure();
        if (depth > _options.MaxBTreeDepth) {
            throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxBTreeDepth),
                depth, _options.MaxBTreeDepth);
        }
        ulong key = PstBinary.NormalizeBid(bid);
        if (!visited.Add(key)) throw new InvalidDataException("The PST subnode tree contains a block cycle.");
        byte[] bytes = ReadBlockPayload(GetBlock(key));
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
                int offset = headerSize + index * entrySize;
                // NIDs remain 32-bit values even though Unicode entries reserve an eight-byte key slot.
                uint nid = PstBinary.UInt32(bytes, offset);
                ulong dataBid = _header.IsUnicode ? PstBinary.UInt64(bytes, offset + 8) : PstBinary.UInt32(bytes, offset + 4);
                ulong subBid = _header.IsUnicode ? PstBinary.UInt64(bytes, offset + 16) : PstBinary.UInt32(bytes, offset + 8);
                if (!result.ContainsKey(nid)) result.Add(nid, new PstSubnodeReference(nid, dataBid, subBid));
            }
        } else {
            int entrySize = _header.IsUnicode ? 16 : 8;
            if (headerSize + checked(count * entrySize) > bytes.Length) {
                throw new InvalidDataException("A PST subnode index block is truncated.");
            }
            for (int index = 0; index < count; index++) {
                int offset = headerSize + index * entrySize;
                ulong child = _header.IsUnicode ? PstBinary.UInt64(bytes, offset + 8) : PstBinary.UInt32(bytes, offset + 4);
                ReadSubnodeBlock(child, result, depth + 1, visited);
            }
        }
        visited.Remove(key);
    }

    private byte[] ReadBlockPayload(PstBlockReference block) {
        if (block.DataLength < 0) throw new InvalidDataException("A PST block has an invalid length.");
        int allocated = PstBinary.Align(
            checked(block.DataLength + _header.BlockTrailerSize), _header.BlockAlignment);
        if (block.Offset < 0 || block.Offset > _stream.Length - allocated) {
            throw new InvalidDataException("A PST block points outside the source stream.");
        }
        byte[] payload = PstBinary.ReadAt(_stream, block.Offset, block.DataLength);
        if (_header.Variant != PstVariant.Unicode4K) return payload;

        int trailerOffset = allocated - _header.BlockTrailerSize;
        byte[] trailer = PstBinary.ReadAt(_stream, block.Offset + trailerOffset, _header.BlockTrailerSize);
        int decodedLength = PstBinary.UInt16(trailer, 18);
        return decodedLength > 0 && decodedLength != payload.Length
            ? PstDeflate.Decode(payload, decodedLength)
            : payload;
    }

    private PstBlockReference GetBlock(ulong normalizedBid) {
        if (!_blocks.TryGetValue(normalizedBid, out PstBlockReference? block)) {
            throw new InvalidDataException(string.Concat("The PST BBT does not contain BID 0x",
                normalizedBid.ToString("X", CultureInfo.InvariantCulture), "."));
        }
        return block;
    }

    private void CountStructure() {
        _cancellationToken.ThrowIfCancellationRequested();
        _visitedStructures++;
        if (_visitedStructures > _options.MaxNodeCount) {
            throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxNodeCount),
                _visitedStructures, _options.MaxNodeCount);
        }
    }
}
