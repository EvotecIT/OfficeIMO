namespace OfficeIMO.Pdf;

internal sealed partial class PdfTrueTypeFontProgram {
    // A composite (CIDFontType2) subset addresses glyphs through CIDToGIDMap, so the embedded
    // program keeps only the painted glyphs, renumbered densely: glyf/loca/hmtx shrink with the
    // glyph set, post carries no names (format 3) and cmap/name keep only a minimal subtable.
    private byte[] BuildCompactSubset(SortedSet<int> glyphs, out int[] glyphOrder) {
        glyphOrder = glyphs.ToArray();
        var newGlyphIds = new Dictionary<int, int>(glyphOrder.Length);
        for (int index = 0; index < glyphOrder.Length; index++) newGlyphIds.Add(glyphOrder[index], index);

        uint[] locaOffsets = ReadLocaOffsets();
        TableRecord glyfTable = GetTable(_tables, "glyf");
        using var glyf = new MemoryStream();
        var offsets = new uint[glyphOrder.Length + 1];
        for (int index = 0; index < glyphOrder.Length; index++) {
            offsets[index] = checked((uint)glyf.Length);
            int glyphId = glyphOrder[index];
            uint start = locaOffsets[glyphId];
            uint end = locaOffsets[glyphId + 1];
            if (end <= start) continue;
            int glyphOffset = checked(glyfTable.Offset + (int)start);
            int glyphLength = checked((int)(end - start));
            EnsureRange(_data, glyphOffset, glyphLength);
            var glyph = new byte[glyphLength];
            Buffer.BlockCopy(_data, glyphOffset, glyph, 0, glyphLength);
            RenumberCompositeComponents(glyph, newGlyphIds);
            glyf.Write(glyph, 0, glyph.Length);
            Pad(glyf, 2);
        }
        offsets[glyphOrder.Length] = checked((uint)glyf.Length);
        bool shortLoca = glyf.Length / 2 <= ushort.MaxValue;

        var tables = new List<SubsetTable> {
            new SubsetTable("glyf", glyf.ToArray()),
            new SubsetTable("loca", BuildCompactLoca(offsets, shortLoca))
        };
        byte[] head = CopyTableData(GetTable(_tables, "head"));
        WriteUInt32(head, 8, 0);
        WriteUInt16(head, 50, shortLoca ? (ushort)0 : (ushort)1);
        tables.Add(new SubsetTable("head", head));

        byte[] hmtx = BuildCompactHorizontalMetrics(glyphOrder, out int numberOfHMetrics);
        byte[] hhea = CopyTableData(GetTable(_tables, "hhea"));
        WriteUInt16(hhea, 34, checked((ushort)numberOfHMetrics));
        tables.Add(new SubsetTable("hhea", hhea));
        tables.Add(new SubsetTable("hmtx", hmtx));

        byte[] maxp = CopyTableData(GetTable(_tables, "maxp"));
        WriteUInt16(maxp, 4, checked((ushort)glyphOrder.Length));
        tables.Add(new SubsetTable("maxp", maxp));

        if (_tables.TryGetValue("post", out TableRecord post) && post.Length >= 32) {
            var compactPost = new byte[32];
            Buffer.BlockCopy(_data, post.Offset, compactPost, 0, 32);
            WriteUInt32(compactPost, 0, 0x00030000);
            tables.Add(new SubsetTable("post", compactPost));
        }

        tables.Add(new SubsetTable("cmap", BuildCompactCmap(newGlyphIds)));
        if (TryBuildCompactName(out byte[] name)) tables.Add(new SubsetTable("name", name));
        else if (_tables.TryGetValue("name", out TableRecord originalName)) tables.Add(new SubsetTable("name", CopyTableData(originalName)));
        foreach (string tag in new[] { "OS/2", "cvt ", "fpgm", "gasp", "prep" }) {
            if (_tables.TryGetValue(tag, out TableRecord record)) tables.Add(new SubsetTable(tag, CopyTableData(record)));
        }

        tables.Sort(static (left, right) => string.CompareOrdinal(left.Tag, right.Tag));
        return BuildTrueTypeFile(tables);
    }

    private static void RenumberCompositeComponents(byte[] glyph, Dictionary<int, int> newGlyphIds) {
        if (glyph.Length < 10 || ReadInt16(glyph, 0) >= 0) return;
        int cursor = 10;
        while (cursor + 4 <= glyph.Length) {
            ushort flags = ReadUInt16(glyph, cursor);
            int component = ReadUInt16(glyph, cursor + 2);
            if (!newGlyphIds.TryGetValue(component, out int renumbered)) {
                throw new InvalidDataException("A TrueType composite glyph references a component outside the subset.");
            }
            WriteUInt16(glyph, cursor + 2, checked((ushort)renumbered));
            cursor += 4;
            cursor += (flags & 0x0001) != 0 ? 4 : 2;
            if ((flags & 0x0008) != 0) cursor += 2;
            else if ((flags & 0x0040) != 0) cursor += 4;
            else if ((flags & 0x0080) != 0) cursor += 8;
            if ((flags & 0x0020) == 0) return;
        }
    }

    private static byte[] BuildCompactLoca(uint[] offsets, bool shortLoca) {
        using var output = new MemoryStream();
        foreach (uint offset in offsets) {
            if (shortLoca) WriteUInt16(output, checked((ushort)(offset / 2)));
            else WriteUInt32(output, offset);
        }
        return output.ToArray();
    }

    private byte[] BuildCompactHorizontalMetrics(int[] glyphOrder, out int numberOfHMetrics) {
        TableRecord hmtx = GetTable(_tables, "hmtx");
        int originalLongMetrics = ReadUInt16(_data, GetTable(_tables, "hhea").Offset + 34);
        if (originalLongMetrics == 0) throw new InvalidDataException("TrueType hhea declares no horizontal metrics.");
        var advances = new ushort[glyphOrder.Length];
        var bearings = new short[glyphOrder.Length];
        for (int index = 0; index < glyphOrder.Length; index++) {
            int glyphId = glyphOrder[index];
            int advanceIndex = Math.Min(glyphId, originalLongMetrics - 1);
            EnsureRange(_data, hmtx.Offset + advanceIndex * 4, 2);
            advances[index] = ReadUInt16(_data, hmtx.Offset + advanceIndex * 4);
            int bearingOffset = glyphId < originalLongMetrics
                ? hmtx.Offset + glyphId * 4 + 2
                : hmtx.Offset + originalLongMetrics * 4 + (glyphId - originalLongMetrics) * 2;
            bearings[index] = bearingOffset + 2 <= hmtx.Offset + hmtx.Length ? ReadInt16(_data, bearingOffset) : (short)0;
        }

        // Trailing glyphs that repeat the last advance store only their side bearings.
        numberOfHMetrics = advances.Length;
        while (numberOfHMetrics > 1 && advances[numberOfHMetrics - 1] == advances[numberOfHMetrics - 2]) numberOfHMetrics--;
        using var output = new MemoryStream();
        for (int index = 0; index < advances.Length; index++) {
            if (index < numberOfHMetrics) WriteUInt16(output, advances[index]);
            WriteUInt16(output, unchecked((ushort)bearings[index]));
        }
        return output.ToArray();
    }

    // A (3,1) format 4 subtable for the subset's BMP characters keeps the program usable
    // on its own (font extraction tools, glyph-name lookups by simple-font companions).
    private byte[] BuildCompactCmap(Dictionary<int, int> newGlyphIds) {
        var mappings = new SortedDictionary<int, int>();
        foreach (KeyValuePair<int, int> entry in _cmap) {
            if (entry.Key <= 0xFFFF && entry.Key != 0xFFFF && newGlyphIds.TryGetValue(entry.Value, out int glyphId) && glyphId > 0)
                mappings[entry.Key] = glyphId;
        }

        var starts = new List<int>();
        var ends = new List<int>();
        var deltas = new List<int>();
        foreach (KeyValuePair<int, int> mapping in mappings) {
            int last = ends.Count - 1;
            if (last >= 0 && ends[last] == mapping.Key - 1 && (deltas[last] & 0xFFFF) == ((mapping.Value - mapping.Key) & 0xFFFF)) {
                ends[last] = mapping.Key;
                continue;
            }
            starts.Add(mapping.Key);
            ends.Add(mapping.Key);
            deltas.Add(mapping.Value - mapping.Key);
        }
        starts.Add(0xFFFF);
        ends.Add(0xFFFF);
        deltas.Add(1);

        int segments = starts.Count;
        int searchRange = 2;
        int entrySelector = 0;
        while (searchRange * 2 <= segments * 2) {
            searchRange *= 2;
            entrySelector++;
        }
        using var subtable = new MemoryStream();
        WriteUInt16(subtable, 4);
        WriteUInt16(subtable, checked((ushort)(16 + segments * 8)));
        WriteUInt16(subtable, 0);
        WriteUInt16(subtable, checked((ushort)(segments * 2)));
        WriteUInt16(subtable, checked((ushort)searchRange));
        WriteUInt16(subtable, checked((ushort)entrySelector));
        WriteUInt16(subtable, checked((ushort)(segments * 2 - searchRange)));
        foreach (int end in ends) WriteUInt16(subtable, (ushort)end);
        WriteUInt16(subtable, 0);
        foreach (int start in starts) WriteUInt16(subtable, (ushort)start);
        foreach (int delta in deltas) WriteUInt16(subtable, unchecked((ushort)delta));
        foreach (int _ in starts) WriteUInt16(subtable, 0);

        using var output = new MemoryStream();
        WriteUInt16(output, 0);
        WriteUInt16(output, 1);
        WriteUInt16(output, 3);
        WriteUInt16(output, 1);
        WriteUInt32(output, 12);
        byte[] data = subtable.ToArray();
        output.Write(data, 0, data.Length);
        return output.ToArray();
    }

    // Family, subfamily, unique, full and PostScript names (Windows Unicode, then Macintosh
    // Roman), without the localized and licensing records of the source font.
    private bool TryBuildCompactName(out byte[] name) {
        name = Array.Empty<byte>();
        if (!_tables.TryGetValue("name", out TableRecord table) || table.Length < 6) return false;
        int count = ReadUInt16(_data, table.Offset + 2);
        int storage = table.Offset + ReadUInt16(_data, table.Offset + 4);
        if (table.Offset + 6 + count * 12 > table.Offset + table.Length) return false;
        var selected = new List<(ushort Platform, ushort Encoding, ushort Language, ushort NameId, byte[] Value)>();
        foreach (ushort nameId in new ushort[] { 1, 2, 3, 4, 6 }) {
            (ushort, ushort, ushort, ushort, byte[])? best = null;
            for (int index = 0; index < count; index++) {
                int record = table.Offset + 6 + index * 12;
                ushort platform = ReadUInt16(_data, record);
                ushort encoding = ReadUInt16(_data, record + 2);
                ushort language = ReadUInt16(_data, record + 4);
                if (ReadUInt16(_data, record + 6) != nameId) continue;
                bool windowsEnglish = platform == 3 && encoding == 1 && language == 0x0409;
                bool macRoman = platform == 1 && encoding == 0 && language == 0;
                if (!windowsEnglish && !(macRoman && best == null)) continue;
                int length = ReadUInt16(_data, record + 8);
                int offset = storage + ReadUInt16(_data, record + 10);
                if (offset < table.Offset || offset + length > table.Offset + table.Length) continue;
                var value = new byte[length];
                Buffer.BlockCopy(_data, offset, value, 0, length);
                best = (platform, encoding, language, nameId, value);
                if (windowsEnglish) break;
            }
            if (best.HasValue) selected.Add(best.Value);
        }
        if (selected.Count == 0) return false;
        selected.Sort(static (left, right) => left.Platform != right.Platform ? left.Platform.CompareTo(right.Platform)
            : left.NameId.CompareTo(right.NameId));

        using var output = new MemoryStream();
        WriteUInt16(output, 0);
        WriteUInt16(output, checked((ushort)selected.Count));
        WriteUInt16(output, checked((ushort)(6 + selected.Count * 12)));
        int stringOffset = 0;
        foreach (var record in selected) {
            WriteUInt16(output, record.Platform);
            WriteUInt16(output, record.Encoding);
            WriteUInt16(output, record.Language);
            WriteUInt16(output, record.NameId);
            WriteUInt16(output, checked((ushort)record.Value.Length));
            WriteUInt16(output, checked((ushort)stringOffset));
            stringOffset += record.Value.Length;
        }
        foreach (var record in selected) output.Write(record.Value, 0, record.Value.Length);
        name = output.ToArray();
        return true;
    }

    private static void WriteUInt16(byte[] data, int offset, ushort value) {
        data[offset] = (byte)(value >> 8);
        data[offset + 1] = (byte)value;
    }
}
