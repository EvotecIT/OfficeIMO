namespace OfficeIMO.Email.Store;

/// <summary>Reads selected values from a DBX indexed record without retaining unrelated property data.</summary>
internal sealed class DbxIndexedRecord {
    private readonly DbxDatabase _database;
    private readonly byte[] _table;
    private readonly long _dataOffset;
    private readonly long _dataLength;

    internal DbxIndexedRecord(DbxDatabase database, byte[] table, long dataOffset, long dataLength) {
        _database = database;
        _table = table;
        _dataOffset = dataOffset;
        _dataLength = dataLength;
    }

    internal uint? ReadUInt32(byte property, CancellationToken cancellationToken) {
        int entry = Find(property);
        if (entry < 0) return null;
        uint value = DbxDatabase.UInt32(_table, entry) >> 8;
        if ((_table[entry] & 0x80) != 0) return value;
        if (EndOfValue(value) - value < 4)
            throw DbxDatabase.Invalid("An indexed integer extends beyond its value.");
        var bytes = new byte[4];
        _database.ReadAt(_dataOffset + value, bytes, 0, bytes.Length, cancellationToken);
        return DbxDatabase.UInt32(bytes, 0);
    }

    internal ulong? ReadUInt64(byte property, CancellationToken cancellationToken) {
        int entry = Find(property);
        if (entry < 0) return null;
        uint value = DbxDatabase.UInt32(_table, entry) >> 8;
        if ((_table[entry] & 0x80) != 0) return value;
        if (EndOfValue(value) - value < 8)
            throw DbxDatabase.Invalid("An indexed 64-bit integer extends beyond its value.");
        var bytes = new byte[8];
        _database.ReadAt(_dataOffset + value, bytes, 0, bytes.Length, cancellationToken);
        return DbxDatabase.UInt32(bytes, 0) | (ulong)DbxDatabase.UInt32(bytes, 4) << 32;
    }

    internal string? ReadString(byte property, CancellationToken cancellationToken,
        long? maximumDecodedBytes = null, long previousDecodedBytes = 0) {
        int entry = Find(property);
        if (entry < 0) return null;
        if ((_table[entry] & 0x80) != 0)
            throw DbxDatabase.Invalid("An indexed string does not reference its data area.");
        uint value = DbxDatabase.UInt32(_table, entry) >> 8;
        long length = EndOfValue(value) - value;
        if (length <= 0) throw DbxDatabase.Invalid("An indexed string starts outside its data area.");
        // Catalog names retain their 4096-byte profile; message strings share the MIME decoded-byte budget.
        long maximum = maximumDecodedBytes.HasValue ? Math.Max(0, maximumDecodedBytes.Value - previousDecodedBytes) : 4096;
        var buffer = new byte[(int)Math.Min(length, Math.Min(4096, Math.Max(1, maximum)))];
        MemoryStream? bytes = null;
        long read = 0;
        try {
            while (read < length && read < maximum) {
                int count = (int)Math.Min(buffer.Length, Math.Min(length - read, maximum - read));
                _database.ReadAt(_dataOffset + value + read, buffer, 0, count, cancellationToken);
                int terminator = Array.IndexOf(buffer, (byte)0, 0, count);
                if (terminator >= 0) {
                    if (bytes == null && terminator == 0) return string.Empty;
                    if (bytes != null) bytes.Write(buffer, 0, terminator);
                    byte[] text = bytes == null ? buffer.AsSpan(0, terminator).ToArray() : bytes.ToArray();
                    if (!MimeTextCodec.TryDecodeTextStrict(text, "windows-1252", out string result))
                        throw DbxDatabase.Invalid("An indexed string cannot be decoded.");
                    return result;
                }
                read += count;
                if (read < length && read < maximum) {
                    bytes ??= new MemoryStream();
                    bytes.Write(buffer, 0, count);
                }
            }
            if (maximumDecodedBytes.HasValue && read < length)
                throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxDecodedPropertyBytesPerItem),
                    checked(previousDecodedBytes + read + 1), maximumDecodedBytes.Value);
            throw DbxDatabase.Invalid("An indexed string has no bounded terminator.");
        } finally {
            bytes?.Dispose();
        }
    }

    private int Find(byte property) {
        int found = -1;
        for (int entry = 0; entry < _table.Length; entry += 4) {
            if ((_table[entry] & 0x7f) != property) continue;
            if (found >= 0) throw DbxDatabase.Invalid("An indexed record repeats a selected property.");
            found = entry;
        }
        return found;
    }

    private long EndOfValue(uint start) {
        long end = _dataLength;
        for (int entry = 0; entry < _table.Length; entry += 4) {
            if ((_table[entry] & 0x80) != 0) continue;
            uint next = DbxDatabase.UInt32(_table, entry) >> 8;
            if (next > start && next < end) end = next;
        }
        return end;
    }
}
