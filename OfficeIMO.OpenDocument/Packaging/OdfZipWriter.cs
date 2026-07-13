namespace OfficeIMO.OpenDocument;

/// <summary>
/// Writes the constrained ZIP shape required by OpenDocument packages without relying on
/// runtime-specific interpretations of <see cref="CompressionLevel.NoCompression"/>.
/// </summary>
internal static class OdfZipWriter {
    private const uint LocalHeaderSignature = 0x04034b50;
    private const uint CentralHeaderSignature = 0x02014b50;
    private const uint EndOfCentralDirectorySignature = 0x06054b50;
    private const ushort Version20 = 20;
    private const ushort Utf8FileNameFlag = 0x0800;
    private const ushort StoredMethod = 0;
    private const ushort DeflateMethod = 8;
    private static readonly uint[] CrcTable = CreateCrcTable();

    internal static byte[] Write(IReadOnlyList<OdfZipWriteEntry> entries, bool deterministic) {
        if (entries == null) throw new ArgumentNullException(nameof(entries));
        if (entries.Count == 0 || !string.Equals(entries[0].Name, "mimetype", StringComparison.Ordinal)) {
            throw new InvalidOperationException("OpenDocument ZIP output must begin with the 'mimetype' entry.");
        }
        if (entries[0].Compress) {
            throw new InvalidOperationException("OpenDocument 'mimetype' must be stored without compression.");
        }
        if (entries.Count > ushort.MaxValue) {
            throw new InvalidOperationException("OpenDocument ZIP output exceeds the non-ZIP64 entry limit.");
        }

        DateTime timestamp = deterministic ? new DateTime(1980, 1, 1, 0, 0, 0, DateTimeKind.Unspecified) : DateTime.Now;
        GetDosTimestamp(timestamp, out ushort dosDate, out ushort dosTime);
        var records = new List<OdfZipRecord>(entries.Count);

        using var output = new MemoryStream();
        using var writer = new BinaryWriter(output, Encoding.UTF8, leaveOpen: true);
        foreach (OdfZipWriteEntry entry in entries) {
            byte[] name = Encoding.UTF8.GetBytes(entry.Name);
            if (name.Length > ushort.MaxValue) {
                throw new InvalidOperationException($"OpenDocument ZIP entry name '{entry.Name}' is too long.");
            }

            byte[] data = entry.Data;
            byte[] payload = entry.Compress ? Deflate(data) : data;
            ushort method = entry.Compress ? DeflateMethod : StoredMethod;
            uint localOffset = ToUInt32(output.Position, "local entry offset");
            uint crc = ComputeCrc32(data);

            WriteLocalHeader(writer, method, dosTime, dosDate, crc, payload.Length, data.Length, name);
            writer.Write(name);
            writer.Write(payload);
            records.Add(new OdfZipRecord(name, method, dosTime, dosDate, crc, payload.Length, data.Length,
                localOffset, entry.Name.EndsWith("/", StringComparison.Ordinal)));
        }

        uint centralOffset = ToUInt32(output.Position, "central directory offset");
        foreach (OdfZipRecord record in records) {
            WriteCentralHeader(writer, record);
            writer.Write(record.Name);
        }
        uint centralSize = ToUInt32(output.Position - centralOffset, "central directory size");
        WriteEndOfCentralDirectory(writer, (ushort)records.Count, centralSize, centralOffset);
        writer.Flush();
        return output.ToArray();
    }

    private static void WriteLocalHeader(BinaryWriter writer, ushort method, ushort time, ushort date, uint crc,
        int compressedLength, int uncompressedLength, byte[] name) {
        writer.Write(LocalHeaderSignature);
        writer.Write(Version20);
        writer.Write(Utf8FileNameFlag);
        writer.Write(method);
        writer.Write(time);
        writer.Write(date);
        writer.Write(crc);
        writer.Write(ToUInt32(compressedLength, "compressed entry size"));
        writer.Write(ToUInt32(uncompressedLength, "uncompressed entry size"));
        writer.Write((ushort)name.Length);
        writer.Write((ushort)0);
    }

    private static void WriteCentralHeader(BinaryWriter writer, OdfZipRecord record) {
        writer.Write(CentralHeaderSignature);
        writer.Write(Version20);
        writer.Write(Version20);
        writer.Write(Utf8FileNameFlag);
        writer.Write(record.Method);
        writer.Write(record.Time);
        writer.Write(record.Date);
        writer.Write(record.Crc);
        writer.Write(ToUInt32(record.CompressedLength, "compressed entry size"));
        writer.Write(ToUInt32(record.UncompressedLength, "uncompressed entry size"));
        writer.Write((ushort)record.Name.Length);
        writer.Write((ushort)0);
        writer.Write((ushort)0);
        writer.Write((ushort)0);
        writer.Write((ushort)0);
        writer.Write(record.IsDirectory ? 0x10u : 0u);
        writer.Write(record.LocalOffset);
    }

    private static void WriteEndOfCentralDirectory(BinaryWriter writer, ushort entryCount, uint centralSize,
        uint centralOffset) {
        writer.Write(EndOfCentralDirectorySignature);
        writer.Write((ushort)0);
        writer.Write((ushort)0);
        writer.Write(entryCount);
        writer.Write(entryCount);
        writer.Write(centralSize);
        writer.Write(centralOffset);
        writer.Write((ushort)0);
    }

    private static byte[] Deflate(byte[] data) {
        using var output = new MemoryStream();
        using (var compressor = new DeflateStream(output, CompressionLevel.Optimal, leaveOpen: true)) {
            compressor.Write(data, 0, data.Length);
        }
        return output.ToArray();
    }

    private static uint ComputeCrc32(byte[] data) {
        uint crc = uint.MaxValue;
        for (int i = 0; i < data.Length; i++) {
            crc = CrcTable[(crc ^ data[i]) & 0xff] ^ (crc >> 8);
        }
        return ~crc;
    }

    private static uint[] CreateCrcTable() {
        var table = new uint[256];
        for (uint index = 0; index < table.Length; index++) {
            uint value = index;
            for (int bit = 0; bit < 8; bit++) {
                value = (value & 1) != 0 ? 0xedb88320u ^ (value >> 1) : value >> 1;
            }
            table[index] = value;
        }
        return table;
    }

    private static void GetDosTimestamp(DateTime value, out ushort date, out ushort time) {
        if (value.Year < 1980) value = new DateTime(1980, 1, 1, 0, 0, 0, DateTimeKind.Unspecified);
        if (value.Year > 2107) value = new DateTime(2107, 12, 31, 23, 59, 58, DateTimeKind.Unspecified);
        date = (ushort)(((value.Year - 1980) << 9) | (value.Month << 5) | value.Day);
        time = (ushort)((value.Hour << 11) | (value.Minute << 5) | (value.Second / 2));
    }

    private static uint ToUInt32(long value, string description) {
        if (value < 0 || value > uint.MaxValue) {
            throw new InvalidOperationException($"OpenDocument ZIP {description} exceeds the non-ZIP64 limit.");
        }
        return (uint)value;
    }

    private sealed class OdfZipRecord {
        internal OdfZipRecord(byte[] name, ushort method, ushort time, ushort date, uint crc, int compressedLength,
            int uncompressedLength, uint localOffset, bool isDirectory) {
            Name = name;
            Method = method;
            Time = time;
            Date = date;
            Crc = crc;
            CompressedLength = compressedLength;
            UncompressedLength = uncompressedLength;
            LocalOffset = localOffset;
            IsDirectory = isDirectory;
        }

        internal byte[] Name { get; }
        internal ushort Method { get; }
        internal ushort Time { get; }
        internal ushort Date { get; }
        internal uint Crc { get; }
        internal int CompressedLength { get; }
        internal int UncompressedLength { get; }
        internal uint LocalOffset { get; }
        internal bool IsDirectory { get; }
    }
}

internal sealed class OdfZipWriteEntry {
    internal OdfZipWriteEntry(string name, byte[] data, bool compress) {
        Name = name ?? throw new ArgumentNullException(nameof(name));
        Data = data ?? throw new ArgumentNullException(nameof(data));
        Compress = compress;
    }

    internal string Name { get; }
    internal byte[] Data { get; }
    internal bool Compress { get; }
}
