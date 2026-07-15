namespace OfficeIMO.Pdf;

internal sealed partial class PdfOpenTypeCffFontProgram {
    private const uint OpenTypeChecksumMagic = 0xB1B0AFBA;

    private static readonly HashSet<string> CompactOpenTypeCffTables = new HashSet<string>(StringComparer.Ordinal) {
        "CFF ",
        "OS/2",
        "cmap",
        "head",
        "hhea",
        "hmtx",
        "maxp",
        "name",
        "post"
    };

    private static readonly HashSet<string> OpenTypeLayoutTables = new HashSet<string>(StringComparer.Ordinal) {
        "BASE",
        "GDEF",
        "GPOS",
        "GSUB",
        "JSTF",
        "MATH"
    };

    internal byte[] BuildCompactOpenTypeFontFile() =>
        BuildCompactOpenTypeFontFilePlan().Data;

    internal PdfOpenTypeCffCompactFontFile BuildCompactOpenTypeFontFilePlan() {
        var tables = new List<SubsetTable>();
        IReadOnlyList<int> usedGlyphIds = GetUsedGlyphIds();
        PdfCffCharStringSubset? charStringSubset = null;
        foreach (KeyValuePair<string, TableRecord> table in _tables.OrderBy(entry => entry.Key, StringComparer.Ordinal)) {
            if (!CompactOpenTypeCffTables.Contains(table.Key)) {
                continue;
            }

            byte[] content = CopyTableData(table.Value);
            if (string.Equals(table.Key, "CFF ", StringComparison.Ordinal)) {
                charStringSubset = PdfCffCharStringSubsetter.Create(content, usedGlyphIds, GlyphCount);
                content = charStringSubset.Data;
            }

            if (string.Equals(table.Key, "head", StringComparison.Ordinal)) {
                WriteUInt32(content, 8, 0);
            }

            tables.Add(new SubsetTable(table.Key, content));
        }

        string[] originalTables = _tables.Keys.OrderBy(tag => tag, StringComparer.Ordinal).ToArray();
        string[] embeddedTables = tables.Select(table => table.Tag).OrderBy(tag => tag, StringComparer.Ordinal).ToArray();
        string[] removedTables = originalTables.Except(embeddedTables, StringComparer.Ordinal).ToArray();
        string[] removedLayoutTables = removedTables.Where(table => OpenTypeLayoutTables.Contains(table)).ToArray();
        if (!ContainsRequiredCompactTables(tables)) {
            return new PdfOpenTypeCffCompactFontFile(
                _data.ToArray(),
                isCompact: false,
                originalTables,
                originalTables,
                Array.Empty<string>(),
                Array.Empty<string>(),
                charStringSubset);
        }

        byte[] compact = BuildOpenTypeFile(tables);
        if (compact.Length < _data.Length || (compact.Length == _data.Length && charStringSubset?.IsSubset == true)) {
            return new PdfOpenTypeCffCompactFontFile(
                compact,
                isCompact: compact.Length < _data.Length,
                originalTables,
                embeddedTables,
                removedTables,
                removedLayoutTables,
                charStringSubset);
        }

        return new PdfOpenTypeCffCompactFontFile(
            _data.ToArray(),
            isCompact: false,
            originalTables,
            originalTables,
            Array.Empty<string>(),
            Array.Empty<string>(),
            charStringSubset: null);
    }

    private static bool ContainsRequiredCompactTables(IReadOnlyList<SubsetTable> tables) {
        var tags = new HashSet<string>(tables.Select(table => table.Tag), StringComparer.Ordinal);
        foreach (string tag in CompactOpenTypeCffTables) {
            if (!tags.Contains(tag)) {
                return false;
            }
        }

        return true;
    }

    private static byte[] BuildOpenTypeFile(IReadOnlyList<SubsetTable> tables) {
        ushort numTables = checked((ushort)tables.Count);
        ushort maxPowerOfTwo = 1;
        ushort entrySelector = 0;
        while (maxPowerOfTwo * 2 <= numTables) {
            maxPowerOfTwo *= 2;
            entrySelector++;
        }

        ushort searchRange = checked((ushort)(maxPowerOfTwo * 16));
        ushort rangeShift = checked((ushort)(numTables * 16 - searchRange));
        using var output = new MemoryStream();
        WriteUInt32(output, OpenTypeCffScalerType);
        WriteUInt16(output, numTables);
        WriteUInt16(output, searchRange);
        WriteUInt16(output, entrySelector);
        WriteUInt16(output, rangeShift);

        long directoryOffset = output.Position;
        for (int index = 0; index < tables.Count; index++) {
            WriteUInt32(output, 0);
            WriteUInt32(output, 0);
            WriteUInt32(output, 0);
            WriteUInt32(output, 0);
        }

        var tableRecords = new List<(SubsetTable Table, uint Checksum, uint Offset)>(tables.Count);
        foreach (SubsetTable table in tables.OrderBy(table => table.Tag, StringComparer.Ordinal)) {
            Pad(output, 4);
            uint offset = checked((uint)output.Position);
            output.Write(table.Data, 0, table.Data.Length);
            uint checksum = CalculateChecksum(table.Data);
            tableRecords.Add((table, checksum, offset));
        }

        byte[] result = output.ToArray();
        for (int index = 0; index < tableRecords.Count; index++) {
            var table = tableRecords[index];
            int offset = checked((int)directoryOffset + index * 16);
            WriteTag(result, offset, table.Table.Tag);
            WriteUInt32(result, offset + 4, table.Checksum);
            WriteUInt32(result, offset + 8, table.Offset);
            WriteUInt32(result, offset + 12, checked((uint)table.Table.Data.Length));
        }

        var headRecord = tableRecords.First(table => string.Equals(table.Table.Tag, "head", StringComparison.Ordinal));
        int headOffset = checked((int)headRecord.Offset);
        WriteUInt32(result, headOffset + 8, 0);
        uint adjustment = unchecked(OpenTypeChecksumMagic - CalculateChecksum(result));
        WriteUInt32(result, headOffset + 8, adjustment);
        return result;
    }

    private byte[] CopyTableData(TableRecord table) {
        var bytes = new byte[table.Length];
        Buffer.BlockCopy(_data, table.Offset, bytes, 0, table.Length);
        return bytes;
    }

    private static uint CalculateChecksum(byte[] data) {
        uint sum = 0;
        int length = Align(data.Length, 4);
        for (int offset = 0; offset < length; offset += 4) {
            uint value =
                ((uint)(offset < data.Length ? data[offset] : 0) << 24) |
                ((uint)(offset + 1 < data.Length ? data[offset + 1] : 0) << 16) |
                ((uint)(offset + 2 < data.Length ? data[offset + 2] : 0) << 8) |
                (uint)(offset + 3 < data.Length ? data[offset + 3] : 0);
            sum = unchecked(sum + value);
        }

        return sum;
    }

    private static int Align(int value, int multiple) =>
        ((value + multiple - 1) / multiple) * multiple;

    private static void Pad(MemoryStream output, int multiple) {
        while (output.Length % multiple != 0) {
            output.WriteByte(0);
        }
    }

    private static void WriteTag(byte[] data, int offset, string tag) {
        if (tag.Length != 4) {
            throw new NotSupportedException("OpenType table tag '" + tag + "' is invalid.");
        }

        byte[] bytes = Encoding.ASCII.GetBytes(tag);
        Buffer.BlockCopy(bytes, 0, data, offset, 4);
    }

    private static void WriteUInt16(MemoryStream output, ushort value) {
        output.WriteByte((byte)(value >> 8));
        output.WriteByte((byte)(value & 0xFF));
    }

    private static void WriteUInt32(MemoryStream output, uint value) {
        output.WriteByte((byte)((value >> 24) & 0xFF));
        output.WriteByte((byte)((value >> 16) & 0xFF));
        output.WriteByte((byte)((value >> 8) & 0xFF));
        output.WriteByte((byte)(value & 0xFF));
    }

    private static void WriteUInt32(byte[] data, int offset, uint value) {
        data[offset] = (byte)((value >> 24) & 0xFF);
        data[offset + 1] = (byte)((value >> 16) & 0xFF);
        data[offset + 2] = (byte)((value >> 8) & 0xFF);
        data[offset + 3] = (byte)(value & 0xFF);
    }

    private readonly struct SubsetTable {
        public SubsetTable(string tag, byte[] data) {
            if (string.IsNullOrWhiteSpace(tag) || tag.Length != 4) {
                throw new ArgumentException("OpenType table tags must be four characters.", nameof(tag));
            }

            Tag = tag;
            Guard.NotNull(data, nameof(data));
            Data = data;
        }

        public string Tag { get; }
        public byte[] Data { get; }
    }
}

internal sealed class PdfOpenTypeCffCompactFontFile {
    public PdfOpenTypeCffCompactFontFile(
        byte[] data,
        bool isCompact,
        IReadOnlyList<string> originalTables,
        IReadOnlyList<string> embeddedTables,
        IReadOnlyList<string> removedTables,
        IReadOnlyList<string> removedLayoutTables,
        PdfCffCharStringSubset? charStringSubset) {
        Data = data ?? throw new ArgumentNullException(nameof(data));
        IsCompact = isCompact;
        OriginalTables = originalTables?.ToArray() ?? throw new ArgumentNullException(nameof(originalTables));
        EmbeddedTables = embeddedTables?.ToArray() ?? throw new ArgumentNullException(nameof(embeddedTables));
        RemovedTables = removedTables?.ToArray() ?? throw new ArgumentNullException(nameof(removedTables));
        RemovedLayoutTables = removedLayoutTables?.ToArray() ?? throw new ArgumentNullException(nameof(removedLayoutTables));
        CharStringSubset = charStringSubset;
    }

    public byte[] Data { get; }
    public bool IsCompact { get; }
    public IReadOnlyList<string> OriginalTables { get; }
    public IReadOnlyList<string> EmbeddedTables { get; }
    public IReadOnlyList<string> RemovedTables { get; }
    public IReadOnlyList<string> RemovedLayoutTables { get; }
    public PdfCffCharStringSubset? CharStringSubset { get; }
}
