using System.Buffers.Binary;

namespace OfficeIMO.Email.Store;

/// <summary>Reads the active index and records of an Outlook Express DBX database.</summary>
internal sealed class DbxDatabase {
    private const int FileHeaderLength = 0x24bc;
    private const int TreeNodeLength = 0x27c;
    private const int TreeNodeHeaderLength = 0x18;
    private const int TreeEntryLength = 12;
    private const int IndexedHeaderLength = 12;
    private const int MaximumTreeDepth = 64;

    private static readonly byte[] Signature = {
        0xcf, 0xad, 0x12, 0xfe, 0xc5, 0xfd, 0x74, 0x6f,
        0x66, 0xe3, 0xd1, 0x11, 0x9a, 0x4e, 0x00, 0xc0,
        0x4f, 0xa3, 0x09, 0xd4, 0x05, 0x00, 0x00, 0x00
    };

    private readonly Stream _source;
    private readonly EmailStoreReaderOptions _options;

    internal DbxDatabase(Stream source, EmailStoreReaderOptions options) {
        _source = source;
        _options = options;
    }

    internal static bool IsCatalogFileName(string path) =>
        string.Equals(Path.GetFileName(path), "Folders.dbx", StringComparison.OrdinalIgnoreCase);

    internal static bool HasSignature(byte[] header, int length, bool folderDatabase = false) {
        if (length < Signature.Length) return false;
        for (int index = 0; index < Signature.Length; index++) {
            byte expected = folderDatabase && index == 4 ? (byte)0xc6 : Signature[index];
            if (header[index] != expected) return false;
        }
        return true;
    }

    internal List<uint> ReadIndex(bool folderDatabase, CancellationToken cancellationToken,
        int indexedItemCount = 0) {
        cancellationToken.ThrowIfCancellationRequested();
        if (_source.Length < FileHeaderLength) throw Invalid("The database header is truncated.");

        var header = new byte[0xf0];
        ReadAt(0, header, 0, header.Length, cancellationToken);
        if (!HasSignature(header, header.Length, folderDatabase))
            throw Invalid("The database type is outside the supported profile.");

        if (!folderDatabase) ValidateLocalMessageDatabase(header, cancellationToken);

        uint expectedCount = UInt32(header, folderDatabase ? 0xcc : 0xc4);
        int maximumCount = folderDatabase ? _options.MaxFolderCount : _options.MaxItemCount;
        string limitName = folderDatabase
            ? nameof(EmailStoreReaderOptions.MaxFolderCount)
            : nameof(EmailStoreReaderOptions.MaxItemCount);
        long totalCount = (long)expectedCount + indexedItemCount;
        if (totalCount > maximumCount)
            throw new EmailStoreLimitExceededException(limitName, totalCount, maximumCount);

        uint root = UInt32(header, folderDatabase ? 0xec : 0xe4);
        var records = new List<uint>();
        if (root != 0) {
            ReadTree(root, parent: 0, depth: 1, expectedCount, records, new HashSet<uint>(),
                new HashSet<uint>(), cancellationToken);
        }
        if (records.Count != expectedCount)
            throw Invalid("The active index does not contain the declared number of records.");
        return records;
    }

    private void ValidateLocalMessageDatabase(byte[] header, CancellationToken cancellationToken) {
        uint length = UInt32(header, 0x1c);
        var fileInfo = new byte[5];
        if (length < fileInfo.Length)
            throw Invalid("The message database is missing its file information.");
        ValidateRange(FileHeaderLength, length);
        ReadAt(FileHeaderLength, fileInfo, 0, fileInfo.Length, cancellationToken);
        if (UInt32(fileInfo, 0) != 1 || fileInfo[4] != 3)
            throw Invalid("The message database is outside the local mail profile.");
    }

    internal DbxIndexedRecord ReadIndexedRecord(uint address, CancellationToken cancellationToken) {
        var header = new byte[IndexedHeaderLength];
        ReadObject(address, header, cancellationToken);
        int propertyCount = header[10];
        if (propertyCount > _options.MaxPropertiesPerItem)
            throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxPropertiesPerItem),
                propertyCount, _options.MaxPropertiesPerItem);
        uint bodyLength = UInt32(header, 4);
        if (bodyLength > _options.MaxDecodedPropertyBytesPerItem)
            throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxDecodedPropertyBytesPerItem),
                bodyLength, _options.MaxDecodedPropertyBytesPerItem);
        int tableLength = propertyCount * 4;
        if (bodyLength < tableLength) throw Invalid("The indexed property table extends past its record.");
        long bodyOffset = (long)address + IndexedHeaderLength;
        ValidateRange(bodyOffset, bodyLength);
        var table = new byte[tableLength];
        ReadAt(bodyOffset, table, 0, table.Length, cancellationToken);
        return new DbxIndexedRecord(this, table, bodyOffset + tableLength, bodyLength - tableLength);
    }

    internal Stream OpenMessage(uint firstFragment, CancellationToken cancellationToken) =>
        new DbxMessageStream(this, firstFragment, _options, cancellationToken);

    internal void ReadObject(uint address, byte[] buffer, CancellationToken cancellationToken) {
        if (address < FileHeaderLength) throw Invalid("A record points into the database header.");
        ReadAt(address, buffer, 0, buffer.Length, cancellationToken);
        if (UInt32(buffer, 0) != address) throw Invalid("A record marker does not match its address.");
    }

    internal void ReadAt(long offset, byte[] buffer, int start, int count,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        ValidateRange(offset, count);
        _source.Position = offset;
        while (count > 0) {
            cancellationToken.ThrowIfCancellationRequested();
            int read = _source.Read(buffer, start, count);
            if (read == 0) throw Invalid("The source ended while reading a declared range.");
            start += read;
            count -= read;
        }
    }

    internal void ValidateRange(long offset, long count) {
        if (offset < 0 || count < 0 || offset > _source.Length - count)
            throw Invalid("A declared range lies outside the source.");
    }

    internal static uint UInt32(byte[] bytes, int offset) =>
        BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(offset, 4));

    internal static InvalidDataException Invalid(string message) =>
        new InvalidDataException("Invalid Outlook Express DBX. " + message);

    private void ReadTree(uint address, uint parent, int depth, uint expectedCount, List<uint> records,
        HashSet<uint> nodesSeen, HashSet<uint> recordsSeen, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        int maximumDepth = Math.Min(_options.MaxBTreeDepth, MaximumTreeDepth);
        if (depth > maximumDepth)
            throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxBTreeDepth), depth, maximumDepth);
        if (nodesSeen.Count >= _options.MaxNodeCount)
            throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxNodeCount),
                nodesSeen.Count + 1L, _options.MaxNodeCount);
        if (!nodesSeen.Add(address)) throw Invalid("The active index repeats a tree node.");

        var node = new byte[TreeNodeLength];
        ReadObject(address, node, cancellationToken);
        if (UInt32(node, 12) != parent) throw Invalid("A tree node has an invalid parent.");
        int entryCount = node[17];
        if (entryCount > (TreeNodeLength - TreeNodeHeaderLength) / TreeEntryLength)
            throw Invalid("A tree node contains too many entries.");

        uint firstChild = UInt32(node, 8);
        if (firstChild != 0)
            ReadTree(firstChild, address, depth + 1, expectedCount, records, nodesSeen, recordsSeen, cancellationToken);
        for (int entry = 0; entry < entryCount; entry++) {
            cancellationToken.ThrowIfCancellationRequested();
            int offset = TreeNodeHeaderLength + entry * TreeEntryLength;
            uint record = UInt32(node, offset);
            if (record < FileHeaderLength || !recordsSeen.Add(record))
                throw Invalid("The active index contains an invalid or repeated record address.");
            ValidateRange(record, IndexedHeaderLength);
            if (records.Count >= expectedCount) throw Invalid("The active index exceeds its declared record count.");
            records.Add(record);

            uint child = UInt32(node, offset + 4);
            if (child != 0)
                ReadTree(child, address, depth + 1, expectedCount, records, nodesSeen, recordsSeen, cancellationToken);
        }
    }
}
