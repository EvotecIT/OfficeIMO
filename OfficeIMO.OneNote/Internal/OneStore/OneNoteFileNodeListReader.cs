namespace OfficeIMO.OneNote;

internal static class OneNoteFileNodeListReader {
    private const ulong HeaderMagic = 0xA4567AB1F5F7F4C4UL;
    private const ulong FooterMagic = 0x8BC215C38233BA4BUL;
    private const int FragmentHeaderLength = 16;
    private const int FragmentTrailerLength = 20;

    public static OneNoteFileNodeList Read(
        Stream stream,
        OneNoteFileChunkReference firstFragment,
        ulong declaredFileLength,
        IReadOnlyDictionary<uint, int> committedNodeCounts,
        OneNoteReaderOptions options) {
        var fragments = new List<OneNoteFileNodeListFragment>();
        var nodes = new List<OneNoteFileNode>();
        var visitedOffsets = new HashSet<ulong>();
        OneNoteFileChunkReference current = firstFragment;
        uint? listId = null;
        int? committedNodeCount = null;
        uint expectedSequence = 0;

        while (!current.IsNil) {
            if (current.IsZero || current.Length < FragmentHeaderLength + FragmentTrailerLength) {
                throw new OneNoteFormatException("ONENOTE_FILE_NODE_FRAGMENT_REFERENCE", "A file-node-list fragment reference is empty or too short.", ToOffset(current.Offset));
            }
            if (fragments.Count >= options.MaxFileNodeListFragments) {
                throw new OneNoteFormatException("ONENOTE_FILE_NODE_FRAGMENT_LIMIT", "The file-node-list fragment limit was exceeded.", ToOffset(current.Offset));
            }
            if (!visitedOffsets.Add(current.Offset)) {
                throw new OneNoteFormatException("ONENOTE_FILE_NODE_FRAGMENT_CYCLE", "The file-node-list fragment chain contains a cycle.", ToOffset(current.Offset));
            }
            ValidateBounds(current.Offset, current.Length, declaredFileLength, "file-node-list fragment");
            if (current.Length > int.MaxValue) {
                throw new OneNoteFormatException("ONENOTE_FILE_NODE_FRAGMENT_SIZE", "A file-node-list fragment is too large to materialize.", ToOffset(current.Offset));
            }

            byte[] data = ReadAt(stream, current.Offset, checked((int)current.Length));
            if (OneNoteBinary.ReadUInt64(data, 0) != HeaderMagic) {
                throw new OneNoteFormatException("ONENOTE_FILE_NODE_HEADER_MAGIC", "The file-node-list fragment header magic is invalid.", ToOffset(current.Offset));
            }
            uint currentListId = OneNoteBinary.ReadUInt32(data, 8);
            uint sequence = OneNoteBinary.ReadUInt32(data, 12);
            if (currentListId < 0x10) {
                throw new OneNoteFormatException("ONENOTE_FILE_NODE_LIST_ID", "The file-node-list identity is below the minimum valid value.", ToOffset(current.Offset + 8));
            }
            if (listId.HasValue && listId.Value != currentListId) {
                throw new OneNoteFormatException("ONENOTE_FILE_NODE_LIST_MISMATCH", "A fragment belongs to a different file-node list.", ToOffset(current.Offset + 8));
            }
            if (!listId.HasValue) {
                if (!committedNodeCounts.TryGetValue(currentListId, out int count)) {
                    throw new OneNoteFormatException("ONENOTE_TRANSACTION_FILE_NODE_LIST", "The transaction log does not declare the referenced file-node list.", ToOffset(current.Offset + 8));
                }
                if (count < 1 || count > options.MaxFileNodes) {
                    throw new OneNoteFormatException("ONENOTE_TRANSACTION_FILE_NODE_COUNT", "The transaction log declares an invalid file-node count.", ToOffset(current.Offset + 8));
                }
                committedNodeCount = count;
            }
            if (sequence != expectedSequence) {
                throw new OneNoteFormatException("ONENOTE_FILE_NODE_SEQUENCE", "The file-node-list fragment sequence is not contiguous.", ToOffset(current.Offset + 12));
            }
            if (OneNoteBinary.ReadUInt64(data, data.Length - 8) != FooterMagic) {
                throw new OneNoteFormatException("ONENOTE_FILE_NODE_FOOTER_MAGIC", "The file-node-list fragment footer magic is invalid.", ToOffset(current.Offset + (ulong)data.Length - 8));
            }

            listId = currentListId;
            int nodeLimit = data.Length - FragmentTrailerLength;
            int remainingNodes = committedNodeCount!.Value - nodes.Count;
            if (remainingNodes <= 0) break;
            var fragmentNodes = ReadNodes(data, FragmentHeaderLength, nodeLimit, current.Offset, declaredFileLength, options, nodes.Count, remainingNodes);
            if (nodes.Count > options.MaxFileNodes - fragmentNodes.Count) {
                throw new OneNoteFormatException("ONENOTE_FILE_NODE_LIMIT", "The file-node limit was exceeded.", ToOffset(current.Offset));
            }
            nodes.AddRange(fragmentNodes);

            OneNoteFileChunkReference next = OneNoteBinary.ReadFileChunkReference64x32(data, data.Length - FragmentTrailerLength);
            fragments.Add(new OneNoteFileNodeListFragment(
                ToOffset(current.Offset),
                data.Length,
                currentListId,
                sequence,
                next,
                fragmentNodes.AsReadOnly()));

            current = next;
            expectedSequence++;
            if (nodes.Count == committedNodeCount.Value) break;
        }

        if (!listId.HasValue) {
            throw new OneNoteFormatException("ONENOTE_FILE_NODE_LIST_EMPTY", "The file-node list contains no fragments.");
        }
        if (nodes.Count != committedNodeCount) {
            throw new OneNoteFormatException("ONENOTE_FILE_NODE_COUNT", "The file-node list ended before its committed transaction-log count was reached.", ToOffset(firstFragment.Offset));
        }
        return new OneNoteFileNodeList(listId.Value, fragments.AsReadOnly(), nodes.AsReadOnly());
    }

    private static List<OneNoteFileNode> ReadNodes(
        byte[] data,
        int start,
        int limit,
        ulong fragmentOffset,
        ulong declaredFileLength,
        OneNoteReaderOptions options,
        int priorNodeCount,
        int maximumNodes) {
        var nodes = new List<OneNoteFileNode>();
        int offset = start;
        while (limit - offset >= 4 && nodes.Count < maximumNodes) {
            if (priorNodeCount + nodes.Count >= options.MaxFileNodes) {
                throw new OneNoteFormatException("ONENOTE_FILE_NODE_LIMIT", "The file-node limit was exceeded.", ToOffset(fragmentOffset + (ulong)offset));
            }
            uint header = OneNoteBinary.ReadUInt32(data, offset);
            ushort rawId = (ushort)(header & 0x3FFU);
            // The final fragment commonly transitions directly into zero-filled padding.
            // File-node identifier zero is not assigned by MS-ONESTORE, so it is a safe
            // padding sentinel when no transaction-log node count has been supplied.
            if (rawId == 0) break;
            int size = (int)((header >> 10) & 0x1FFFU);
            byte stpFormat = (byte)((header >> 23) & 0x03U);
            byte cbFormat = (byte)((header >> 25) & 0x03U);
            byte rawBaseType = (byte)((header >> 27) & 0x0FU);
            if ((header & 0x80000000U) == 0) {
                throw new OneNoteFormatException("ONENOTE_FILE_NODE_RESERVED_BIT", "The required file-node reserved bit is not set.", ToOffset(fragmentOffset + (ulong)offset));
            }
            if (size < 4 || size > limit - offset) {
                throw new OneNoteFormatException("ONENOTE_FILE_NODE_SIZE", "The file-node size is invalid or crosses the fragment trailer.", ToOffset(fragmentOffset + (ulong)offset));
            }
            if (rawBaseType > 2) {
                throw new OneNoteFormatException("ONENOTE_FILE_NODE_BASE_TYPE", "The file-node base type is invalid.", ToOffset(fragmentOffset + (ulong)offset));
            }

            int dataLength = size - 4;
            var encodedData = new byte[dataLength];
            if (dataLength > 0) Buffer.BlockCopy(data, offset + 4, encodedData, 0, dataLength);
            OneNoteFileNodeChunkReference? chunkReference = null;
            var baseType = (OneNoteFileNodeBaseType)rawBaseType;
            if (baseType != OneNoteFileNodeBaseType.Inline) {
                chunkReference = ReadChunkReference(encodedData, stpFormat, cbFormat, fragmentOffset + (ulong)offset + 4);
                if (!chunkReference.IsNil && !chunkReference.IsZero) {
                    ValidateBounds(chunkReference.Offset, chunkReference.Length, declaredFileLength, "file-node chunk reference");
                }
            } else if (cbFormat != 0) {
                throw new OneNoteFormatException("ONENOTE_INLINE_CB_FORMAT", "An inline file node has a nonzero byte-count format.", ToOffset(fragmentOffset + (ulong)offset));
            }

            nodes.Add(new OneNoteFileNode(
                rawId,
                size,
                stpFormat,
                cbFormat,
                baseType,
                ToOffset(fragmentOffset + (ulong)offset),
                chunkReference,
                encodedData));
            offset += size;

            if (rawId == (ushort)OneNoteFileNodeId.ChunkTerminator) break;
        }
        return nodes;
    }

    private static OneNoteFileNodeChunkReference ReadChunkReference(byte[] data, byte stpFormat, byte cbFormat, ulong absoluteOffset) {
        int stpBytes = stpFormat == 0 ? 8 : stpFormat == 2 ? 2 : 4;
        int cbBytes = cbFormat == 0 ? 4 : cbFormat == 1 ? 8 : cbFormat == 2 ? 1 : 2;
        OneNoteBinary.EnsureRange(data, 0, stpBytes + cbBytes);
        ulong rawOffset = ReadUnsigned(data, 0, stpBytes);
        ulong rawLength = ReadUnsigned(data, stpBytes, cbBytes);
        bool compressedOffset = stpFormat >= 2;
        bool compressedLength = cbFormat >= 2;
        bool isNil = rawOffset == MaxValue(stpBytes) && rawLength == 0;
        ulong offset = compressedOffset && !isNil ? CheckedMultiplyByEight(rawOffset, absoluteOffset) : rawOffset;
        ulong length = compressedLength ? CheckedMultiplyByEight(rawLength, absoluteOffset + (ulong)stpBytes) : rawLength;
        return new OneNoteFileNodeChunkReference(offset, length, isNil, stpBytes + cbBytes);
    }

    private static ulong ReadUnsigned(byte[] data, int offset, int length) {
        OneNoteBinary.EnsureRange(data, offset, length);
        ulong value = 0;
        for (int index = 0; index < length; index++) value |= (ulong)data[offset + index] << (index * 8);
        return value;
    }

    private static ulong MaxValue(int byteCount) => byteCount == 8 ? ulong.MaxValue : (1UL << (byteCount * 8)) - 1;

    private static ulong CheckedMultiplyByEight(ulong value, ulong offset) {
        if (value > ulong.MaxValue / 8) {
            throw new OneNoteFormatException("ONENOTE_COMPRESSED_REFERENCE_OVERFLOW", "A compressed chunk reference overflows its decoded range.", ToOffset(offset));
        }
        return value * 8;
    }

    private static byte[] ReadAt(Stream stream, ulong offset, int length) {
        long position = ToOffset(offset);
        stream.Position = position;
        var buffer = new byte[length];
        int total = 0;
        while (total < buffer.Length) {
            int read = stream.Read(buffer, total, buffer.Length - total);
            if (read <= 0) {
                throw new OneNoteFormatException("ONENOTE_TRUNCATED_STRUCTURE", "The OneNote file ended while reading a referenced structure.", position + total);
            }
            total += read;
        }
        return buffer;
    }

    private static void ValidateBounds(ulong offset, ulong length, ulong fileLength, string name) {
        if (offset > fileLength || length > fileLength - offset) {
            throw new OneNoteFormatException("ONENOTE_CHUNK_REFERENCE_BOUNDS", "The " + name + " lies outside the declared file length.", ToOffset(offset));
        }
    }

    private static long ToOffset(ulong offset) {
        if (offset > long.MaxValue) throw new OneNoteFormatException("ONENOTE_OFFSET_RANGE", "A OneNote file offset exceeds the supported signed range.");
        return (long)offset;
    }
}
