namespace OfficeIMO.OneNote;

/// <summary>Resolves committed file-node-list sizes from the MS-ONESTORE transaction log.</summary>
internal static class OneNoteTransactionLogReader {
    private const int TransactionEntryLength = 8;
    private const int NextFragmentLength = 12;

    internal static IReadOnlyDictionary<uint, int> Read(
        Stream stream,
        OneNoteFileHeader header,
        OneNoteReaderOptions options) {
        if (!header.TransactionLog.HasValue || !header.TransactionCount.HasValue || !header.ExpectedFileLength.HasValue) {
            throw new OneNoteFormatException("ONENOTE_TRANSACTION_LOG_HEADER", "The revision-store header does not expose a complete transaction log reference.");
        }

        OneNoteFileChunkReference current = header.TransactionLog.Value;
        uint committedTransactionCount = header.TransactionCount.Value;
        uint completedTransactions = 0;
        int fragmentCount = 0;
        int entryCount = 0;
        uint runningCrc = 0;
        var visitedOffsets = new HashSet<ulong>();
        var nodeCounts = new Dictionary<uint, int>();

        while (completedTransactions < committedTransactionCount) {
            if (current.IsNil || current.IsZero || current.Length < TransactionEntryLength + NextFragmentLength) {
                throw new OneNoteFormatException("ONENOTE_TRANSACTION_LOG_TRUNCATED", "The transaction log ended before all committed transactions were found.", ToOffset(current.Offset));
            }
            if (++fragmentCount > options.MaxTransactionLogFragments) {
                throw new OneNoteFormatException("ONENOTE_TRANSACTION_FRAGMENT_LIMIT", "The transaction-log fragment limit was exceeded.", ToOffset(current.Offset));
            }
            if (!visitedOffsets.Add(current.Offset)) {
                throw new OneNoteFormatException("ONENOTE_TRANSACTION_FRAGMENT_CYCLE", "The transaction-log fragment chain contains a cycle.", ToOffset(current.Offset));
            }
            ValidateBounds(current, header.ExpectedFileLength.Value);
            if (current.Length > int.MaxValue) {
                throw new OneNoteFormatException(
                    "ONENOTE_TRANSACTION_FRAGMENT_SIZE",
                    "A transaction-log fragment is too large to materialize.",
                    ToOffset(current.Offset));
            }
            byte[] data = ReadAt(stream, current.Offset, checked((int)current.Length));

            int entryBytes = ((data.Length - NextFragmentLength) / TransactionEntryLength) * TransactionEntryLength;
            int offset = 0;
            while (offset < entryBytes && completedTransactions < committedTransactionCount) {
                if (++entryCount > options.MaxTransactionEntries) {
                    throw new OneNoteFormatException("ONENOTE_TRANSACTION_ENTRY_LIMIT", "The transaction-entry limit was exceeded.", ToOffset(current.Offset + (ulong)offset));
                }
                uint sourceId = OneNoteBinary.ReadUInt32(data, offset);
                uint value = OneNoteBinary.ReadUInt32(data, offset + 4);

                if (sourceId == 1) {
                    if (options.ValidateTransactionChecksums && value != runningCrc) {
                        throw new OneNoteFormatException(
                            "ONENOTE_TRANSACTION_CHECKSUM",
                            "A committed transaction has an invalid sentinel checksum.",
                            ToOffset(current.Offset + (ulong)offset + 4));
                    }
                    completedTransactions++;
                    runningCrc = OneNoteCrc32.Continue(runningCrc, data, offset, TransactionEntryLength, header.FileKind);
                    offset += TransactionEntryLength;
                    continue;
                }
                if (sourceId < 0x10) {
                    throw new OneNoteFormatException("ONENOTE_TRANSACTION_SOURCE_ID", "A transaction entry contains an invalid file-node-list identity.", ToOffset(current.Offset + (ulong)offset));
                }
                if (value > int.MaxValue || value > (uint)options.MaxFileNodes) {
                    throw new OneNoteFormatException("ONENOTE_TRANSACTION_FILE_NODE_COUNT", "A transaction entry contains an invalid file-node count.", ToOffset(current.Offset + (ulong)offset + 4));
                }
                int newCount = (int)value;
                if (nodeCounts.TryGetValue(sourceId, out int previousCount) && newCount <= previousCount) {
                    throw new OneNoteFormatException("ONENOTE_TRANSACTION_FILE_NODE_SEQUENCE", "A transaction entry does not increase its file-node-list count.", ToOffset(current.Offset + (ulong)offset + 4));
                }
                nodeCounts[sourceId] = newCount;
                runningCrc = OneNoteCrc32.Continue(runningCrc, data, offset, TransactionEntryLength, header.FileKind);
                offset += TransactionEntryLength;
            }

            if (completedTransactions >= committedTransactionCount) break;
            current = OneNoteBinary.ReadFileChunkReference64x32(data, entryBytes);
        }

        if (nodeCounts.Count == 0) {
            throw new OneNoteFormatException("ONENOTE_TRANSACTION_LOG_EMPTY", "The committed transaction log declares no file-node lists.");
        }
        return nodeCounts;
    }

    private static void ValidateBounds(OneNoteFileChunkReference reference, ulong declaredFileLength) {
        if (reference.Offset > declaredFileLength || reference.Length > declaredFileLength - reference.Offset) {
            throw new OneNoteFormatException("ONENOTE_TRANSACTION_FRAGMENT_BOUNDS", "A transaction-log fragment lies outside the declared file length.", ToOffset(reference.Offset));
        }
    }

    private static byte[] ReadAt(Stream stream, ulong offset, int length) {
        stream.Position = ToOffset(offset);
        var data = new byte[length];
        int total = 0;
        while (total < data.Length) {
            int read = stream.Read(data, total, data.Length - total);
            if (read <= 0) throw new OneNoteFormatException("ONENOTE_TRANSACTION_FRAGMENT_TRUNCATED", "The file ended while reading a transaction-log fragment.", ToOffset(offset) + total);
            total += read;
        }
        return data;
    }

    private static long ToOffset(ulong value) {
        if (value > long.MaxValue) throw new OneNoteFormatException("ONENOTE_OFFSET_RANGE", "A OneNote file offset exceeds the supported stream range.");
        return (long)value;
    }
}
