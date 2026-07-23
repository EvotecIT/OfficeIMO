namespace OfficeIMO.Pdf;

internal sealed partial class PdfTrueTypeFontProgram {
    private const uint TrueTypeChecksumMagic = 0xB1B0AFBA;

    internal byte[] BuildSubsetFontFile() {
        var glyphs = new SortedSet<int>(GetUsedGlyphIds());
        if (!_tables.TryGetValue("glyf", out TableRecord originalGlyf) ||
            !_tables.TryGetValue("loca", out TableRecord originalLoca)) {
            return _data.ToArray();
        }

        glyphs.Add(0);
        AddCompositeGlyphDependencies(glyphs);

        uint[] locaOffsets = ReadLocaOffsets();
        byte[] glyf = BuildSubsetGlyfTable(glyphs, locaOffsets, out uint[] subsetLocaOffsets);
        byte[] loca = BuildLocaTable(subsetLocaOffsets);

        var tables = new List<SubsetTable>();
        foreach (var table in _tables.OrderBy(entry => entry.Key, StringComparer.Ordinal)) {
            string tag = table.Key;
            if (string.Equals(tag, "DSIG", StringComparison.Ordinal)) {
                continue;
            }

            byte[] content;
            if (string.Equals(tag, "glyf", StringComparison.Ordinal)) {
                content = glyf;
            } else if (string.Equals(tag, "loca", StringComparison.Ordinal)) {
                content = loca;
            } else if (string.Equals(tag, "head", StringComparison.Ordinal)) {
                content = CopyTableData(table.Value);
                WriteUInt32(content, 8, 0);
            } else {
                content = CopyTableData(table.Value);
            }

            tables.Add(new SubsetTable(tag, content));
        }

        if (tables.Count == _tables.Count &&
            glyf.Length >= originalGlyf.Length &&
            loca.Length >= originalLoca.Length) {
            return _data.ToArray();
        }

        return BuildTrueTypeFile(tables);
    }

    private void AddCompositeGlyphDependencies(SortedSet<int> glyphs) {
        uint[] locaOffsets = ReadLocaOffsets();
        TableRecord glyfTable = GetTable(_tables, "glyf");
        var queue = new Queue<int>(glyphs);
        var visited = new HashSet<int>();

        while (queue.Count > 0) {
            int glyphId = queue.Dequeue();
            if (!visited.Add(glyphId) || glyphId < 0 || glyphId >= GlyphCount) {
                continue;
            }

            foreach (int component in EnumerateCompositeGlyphComponents(glyfTable, locaOffsets, glyphId)) {
                if (glyphs.Add(component)) {
                    queue.Enqueue(component);
                }
            }
        }
    }

    private IEnumerable<int> EnumerateCompositeGlyphComponents(TableRecord glyfTable, uint[] locaOffsets, int glyphId) {
        uint start = locaOffsets[glyphId];
        uint end = locaOffsets[glyphId + 1];
        if (end <= start) {
            yield break;
        }

        int glyphOffset = checked(glyfTable.Offset + (int)start);
        int glyphLength = checked((int)(end - start));
        EnsureRange(_data, glyphOffset, glyphLength);
        if (glyphLength < 10 || ReadInt16(_data, glyphOffset) >= 0) {
            yield break;
        }

        int cursor = glyphOffset + 10;
        int glyphEnd = glyphOffset + glyphLength;
        while (cursor + 4 <= glyphEnd) {
            ushort flags = ReadUInt16(_data, cursor);
            int componentGlyphId = ReadUInt16(_data, cursor + 2);
            yield return componentGlyphId;
            cursor += 4;

            cursor += (flags & 0x0001) != 0 ? 4 : 2;
            if ((flags & 0x0008) != 0) {
                cursor += 2;
            } else if ((flags & 0x0040) != 0) {
                cursor += 4;
            } else if ((flags & 0x0080) != 0) {
                cursor += 8;
            }

            if ((flags & 0x0020) == 0) {
                yield break;
            }
        }
    }

    private uint[] ReadLocaOffsets() {
        TableRecord head = GetTable(_tables, "head");
        TableRecord loca = GetTable(_tables, "loca");
        int indexToLocFormat = ReadInt16(_data, head.Offset + 50);
        var offsets = new uint[GlyphCount + 1];

        if (indexToLocFormat == 0) {
            EnsureRange(_data, loca.Offset, checked((GlyphCount + 1) * 2));
            for (int index = 0; index < offsets.Length; index++) {
                offsets[index] = (uint)(ReadUInt16(_data, loca.Offset + index * 2) * 2);
            }
        } else if (indexToLocFormat == 1) {
            EnsureRange(_data, loca.Offset, checked((GlyphCount + 1) * 4));
            for (int index = 0; index < offsets.Length; index++) {
                offsets[index] = ReadUInt32(_data, loca.Offset + index * 4);
            }
        } else {
            throw new NotSupportedException("TrueType font has an unsupported loca index format.");
        }

        return offsets;
    }

    private byte[] BuildSubsetGlyfTable(SortedSet<int> glyphs, uint[] locaOffsets, out uint[] subsetLocaOffsets) {
        TableRecord glyfTable = GetTable(_tables, "glyf");
        subsetLocaOffsets = new uint[GlyphCount + 1];
        using var output = new MemoryStream();

        for (int glyphId = 0; glyphId < GlyphCount; glyphId++) {
            subsetLocaOffsets[glyphId] = checked((uint)output.Length);
            uint start = locaOffsets[glyphId];
            uint end = locaOffsets[glyphId + 1];
            if (!glyphs.Contains(glyphId) || end <= start) {
                continue;
            }

            int glyphOffset = checked(glyfTable.Offset + (int)start);
            int glyphLength = checked((int)(end - start));
            EnsureRange(_data, glyphOffset, glyphLength);
            output.Write(_data, glyphOffset, glyphLength);
            Pad(output, 4);
        }

        subsetLocaOffsets[GlyphCount] = checked((uint)output.Length);
        return output.ToArray();
    }

    private byte[] BuildLocaTable(uint[] offsets) {
        int indexToLocFormat = ReadInt16(_data, GetTable(_tables, "head").Offset + 50);
        using var output = new MemoryStream();
        if (indexToLocFormat == 0) {
            foreach (uint offset in offsets) {
                if ((offset & 1) != 0 || offset / 2 > ushort.MaxValue) {
                    throw new NotSupportedException("TrueType short loca offsets exceed the supported subset range.");
                }

                WriteUInt16(output, (ushort)(offset / 2));
            }
        } else {
            foreach (uint offset in offsets) {
                WriteUInt32(output, offset);
            }
        }

        return output.ToArray();
    }

    private static byte[] BuildTrueTypeFile(IReadOnlyList<SubsetTable> tables) {
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
        WriteUInt32(output, 0x00010000);
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
        foreach (SubsetTable table in tables) {
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
        uint adjustment = unchecked(TrueTypeChecksumMagic - CalculateChecksum(result));
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
            throw new NotSupportedException("TrueType table tag '" + tag + "' is invalid.");
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
                throw new ArgumentException("TrueType table tags must be four characters.", nameof(tag));
            }

            Tag = tag;
            Data = data ?? throw new ArgumentNullException(nameof(data));
        }

        public string Tag { get; }
        public byte[] Data { get; }
    }
}
