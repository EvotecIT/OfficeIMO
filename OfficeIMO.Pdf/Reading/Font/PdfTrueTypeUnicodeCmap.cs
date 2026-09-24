using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>
/// Gives embedded simple TrueType programs a Unicode cmap for drawing-scene glyph lookup.
/// </summary>
/// <remarks>
/// A PDF simple TrueType font selects glyphs by character code through a (3,0) symbolic or (1,0)
/// Macintosh cmap (ISO 32000-1, 9.6.6.4). Drawing scenes carry Unicode text, so a program without a
/// Unicode subtable cannot resolve the glyph; a program with one can resolve a different glyph.
/// This helper maps each code's decoded text to the glyph that code paints and rebuilds the cmap.
/// </remarks>
internal static class PdfTrueTypeUnicodeCmap {
    private const uint ChecksumMagic = 0xB1B0AFBA;
    private const int SymbolicFlag = 1 << 2;

    /// <summary>Returns a drawing program with a synthesized Unicode cmap, or null when none is needed or possible.</summary>
    /// <param name="font">The PDF font resource whose encoding and ToUnicode map describe the program codes.</param>
    /// <param name="program">The decoded embedded TrueType program.</param>
    /// <param name="cidToGlyphMap">A CIDFontType2 CIDToGIDMap (empty for identity), or null for simple fonts.</param>
    internal static PdfDrawingFontProgram? TryCreate(PdfFontResource font, byte[] program, byte[]? cidToGlyphMap) {
        if (!TryReadTables(program, out List<(string Tag, int Offset, int Length)> tables)) return null;
        foreach (string required in new[] { "glyf", "head", "hhea", "hmtx", "loca", "maxp" }) {
            if (!tables.Exists(table => string.Equals(table.Tag, required, StringComparison.Ordinal))) return null;
        }
        int cmapIndex = tables.FindIndex(static table => string.Equals(table.Tag, "cmap", StringComparison.Ordinal));
        int? symbolic = null;
        int? macintosh = null;
        if (cmapIndex >= 0 &&
            !TryReadCodeSubtables(program, tables[cmapIndex].Offset, tables[cmapIndex].Length, out symbolic, out macintosh, out _)) return null;
        int glyphCount = ReadGlyphCount(program, tables);
        Func<int, bool> isEmptyGlyph = CreateEmptyGlyphTest(program, tables, glyphCount);

        SortedDictionary<int, int>? mappings;
        Func<int, int> glyphForCode;
        if (cidToGlyphMap != null) {
            // CIDFontType2 glyphs are selected only through CIDToGIDMap; an embedded Unicode cmap need
            // not describe what the page paints, so always derive the drawing cmap from the PDF.
            mappings = CreateCidMappings(font, cidToGlyphMap, glyphCount, isEmptyGlyph);
            glyphForCode = cid => GlyphForCid(cidToGlyphMap, cid, glyphCount);
        } else if (string.Equals(font.FontSubtype, "TrueType", StringComparison.Ordinal) &&
            string.Equals(font.EmbeddedProgramSubtype, "TrueType", StringComparison.Ordinal) &&
            (symbolic.HasValue || macintosh.HasValue)) {
            // PDF character codes select through the symbolic or Macintosh table. A Unicode
            // subtable can name a different glyph than the one that the PDF actually paints.
            mappings = CreateSimpleMappings(font, program, symbolic, macintosh, glyphCount, isEmptyGlyph);
            bool isSymbolic = ((font.FontDescriptorFlags ?? SymbolicFlag) & SymbolicFlag) != 0;
            Func<byte, string> decodeEncoding = ResourceResolver.CreateSimpleEncodingDecoder(font);
            glyphForCode = code => code is < 0 or > 255 ? 0 : ResolveGlyph(program, (byte)code, symbolic, macintosh, isSymbolic, decodeEncoding);
        } else {
            return null;
        }
        if (mappings == null || mappings.Count == 0) return null;
        byte[]? rebuilt = Rebuild(program, mappings);
        return rebuilt == null ? null : new PdfDrawingFontProgram(rebuilt, mappings, glyphForCode, isEmptyGlyph);
    }

    /// <summary>Returns the program with additional Unicode-to-glyph mappings, keeping existing entries.</summary>
    internal static byte[]? TryAddMappings(PdfDrawingFontProgram source, IReadOnlyDictionary<int, int> additions) {
        var merged = new SortedDictionary<int, int>();
        foreach (KeyValuePair<int, int> existing in source.UnicodeGlyphs) merged.Add(existing.Key, existing.Value);
        foreach (KeyValuePair<int, int> addition in additions) {
            if (!merged.ContainsKey(addition.Key)) merged.Add(addition.Key, addition.Value);
        }
        return merged.Count == source.UnicodeGlyphs.Count ? null : Rebuild(source.Program, merged);
    }

    private static int GlyphForCid(byte[] cidToGlyphMap, int cid, int glyphCount) {
        int glyph = cidToGlyphMap.Length == 0
            ? cid
            : cid >= 0 && 2 * cid + 1 < cidToGlyphMap.Length ? (cidToGlyphMap[2 * cid] << 8) | cidToGlyphMap[2 * cid + 1] : 0;
        return glyph > 0 && glyph < glyphCount ? glyph : 0;
    }

    private static byte[]? Rebuild(byte[] program, SortedDictionary<int, int> mappings) {
        if (!TryReadTables(program, out List<(string Tag, int Offset, int Length)> tables)) return null;
        var bodies = new List<(string Tag, byte[] Body)>(tables.Count + 1);
        foreach (var table in tables) {
            if (string.Equals(table.Tag, "cmap", StringComparison.Ordinal)) continue;
            var body = new byte[table.Length];
            Buffer.BlockCopy(program, table.Offset, body, 0, table.Length);
            bodies.Add((table.Tag, body));
        }
        bodies.Add(("cmap", BuildUnicodeCmap(mappings)));
        bodies.Sort(static (left, right) => string.CompareOrdinal(left.Tag, right.Tag));
        byte[] rebuilt = Assemble(bodies);
        return OfficeTrueTypeFont.TryLoad(rebuilt) == null ? null : rebuilt;
    }

    private static SortedDictionary<int, int> CreateSimpleMappings(PdfFontResource font, byte[] program, int? symbolic, int? macintosh,
        int glyphCount, Func<int, bool> isEmptyGlyph) {
        bool isSymbolic = ((font.FontDescriptorFlags ?? SymbolicFlag) & SymbolicFlag) != 0;
        Func<byte[], int, string> decodeText = ResourceResolver.CreateBudgetedDecoder(font);
        Func<byte, string> decodeEncoding = ResourceResolver.CreateSimpleEncodingDecoder(font);
        var mappings = new SortedDictionary<int, int>();
        int fallbackGlyph = 0;
        for (int code = 0; code < 256; code++) {
            int glyph = ResolveGlyph(program, (byte)code, symbolic, macintosh, isSymbolic, decodeEncoding);
            if (glyph <= 0 || glyph >= glyphCount) continue;
            if (fallbackGlyph == 0 && !isEmptyGlyph(glyph)) fallbackGlyph = glyph;
            string text;
            try {
                text = decodeText(new[] { (byte)code }, 8);
            } catch (PdfReadLimitException) {
                continue;
            }
            if (!TryGetSingleScalar(text, out int scalar) || mappings.ContainsKey(scalar) ||
                IsWhiteSpaceScalar(scalar) && !isEmptyGlyph(glyph)) continue;
            mappings.Add(scalar, glyph);
        }
        // A simple subset may contain only clusters or inked codes mapped to whitespace.
        // A private-use seed registers the face; the painted-run mapper assigns exact aliases.
        if (mappings.Count == 0 && fallbackGlyph > 0) mappings.Add(0xE000, fallbackGlyph);
        return mappings;
    }

    // A CIDFontType2 font selects glyphs by CID through CIDToGIDMap. With an Identity encoding the
    // two-byte code is the CID, and ToUnicode gives the text that CID paints.
    private static SortedDictionary<int, int>? CreateCidMappings(PdfFontResource font, byte[] cidToGlyphMap, int glyphCount,
        Func<int, bool> isEmptyGlyph) {
        if (font.CMap == null ||
            !string.Equals(font.Encoding, "Identity-H", StringComparison.Ordinal) &&
            !string.Equals(font.Encoding, "Identity-V", StringComparison.Ordinal)) return null;
        var entries = new List<(int Cid, int Scalar)>();
        int fallbackGlyph = 0;
        foreach (KeyValuePair<string, string> mapping in font.CMap.Mappings) {
            if (mapping.Key.Length != 4 ||
                !int.TryParse(mapping.Key, System.Globalization.NumberStyles.HexNumber, System.Globalization.CultureInfo.InvariantCulture, out int cid)) continue;
            int paintedGlyph = GlyphForCid(cidToGlyphMap, cid, glyphCount);
            if (fallbackGlyph == 0 && paintedGlyph > 0 && !isEmptyGlyph(paintedGlyph)) fallbackGlyph = paintedGlyph;
            if (!TryGetSingleScalar(mapping.Value, out int scalar)) continue;
            entries.Add((cid, scalar));
        }
        entries.Sort(static (left, right) => left.Cid.CompareTo(right.Cid));

        var mappings = new SortedDictionary<int, int>();
        foreach (var entry in entries) {
            int glyph = GlyphForCid(cidToGlyphMap, entry.Cid, glyphCount);
            // Several glyphs can decode to one letter (shaped contextual forms). The first keeps the
            // base letter; drawings add presentation-form entries for the forms a page paints.
            // Producers map colour-font layers and other inked glyphs to a space; whitespace must
            // only select an empty glyph, or every space in the font would paint that ink.
            if (glyph <= 0 || mappings.ContainsKey(entry.Scalar) ||
                IsWhiteSpaceScalar(entry.Scalar) && !isEmptyGlyph(glyph)) continue;
            mappings.Add(entry.Scalar, glyph);
        }
        // A subset can contain only clusters, undecoded glyphs, or inked whitespace. Seed one private-use entry so the
        // drawing font is registered; the painted-glyph mapper assigns aliases for each run.
        if (mappings.Count == 0 && fallbackGlyph > 0) mappings.Add(0xE000, fallbackGlyph);
        return mappings;
    }

    private static int ResolveGlyph(byte[] data, byte code, int? symbolic, int? macintosh, bool isSymbolic,
        Func<byte, string> decodeEncoding) {
        if (!isSymbolic && macintosh.HasValue) {
            // A nonsymbolic font selects its PDF-encoded name through Mac Roman before any
            // symbolic table, even when the same program also carries a (3,0) subtable.
            string named = decodeEncoding(code);
            if (named.Length == 1 && PdfMacRomanEncoding.TryEncode(named[0], out byte macCode)) {
                int glyph = LookupSubtable(data, macintosh.Value, macCode);
                if (glyph > 0) return glyph;
            }
        }
        if (symbolic.HasValue) {
            foreach (int prefix in new[] { 0x0000, 0xF000, 0xF100, 0xF200 }) {
                int glyph = LookupSubtable(data, symbolic.Value, prefix | code);
                if (glyph > 0) return glyph;
            }
        }
        if (!macintosh.HasValue) return 0;
        return LookupSubtable(data, macintosh.Value, code);
    }

    private static bool TryGetSingleScalar(string text, out int scalar) {
        scalar = 0;
        if (text.Length == 1 && !char.IsSurrogate(text[0])) scalar = text[0];
        else if (text.Length == 2 && char.IsSurrogatePair(text[0], text[1])) scalar = char.ConvertToUtf32(text[0], text[1]);
        else return false;
        // Controls, the replacement character and noncharacters are placeholders, not painted text.
        return scalar >= 0x20 && scalar != 0xFFFD && !(scalar >= 0x7F && scalar < 0xA0) &&
            (scalar & 0xFFFE) != 0xFFFE && !(scalar >= 0xFDD0 && scalar <= 0xFDEF);
    }

    private static bool TryReadTables(byte[] data, out List<(string Tag, int Offset, int Length)> tables) {
        tables = new List<(string Tag, int Offset, int Length)>();
        if (data.Length < 12) return false;
        uint version = ReadUInt32(data, 0);
        if (version != 0x00010000 && version != 0x74727565) return false;
        int count = ReadUInt16(data, 4);
        if (count == 0 || 12 + count * 16 > data.Length) return false;
        for (int index = 0; index < count; index++) {
            int record = 12 + index * 16;
            uint offset = ReadUInt32(data, record + 8);
            uint length = ReadUInt32(data, record + 12);
            if (offset > (uint)data.Length || length > (uint)data.Length - offset) return false;
            tables.Add((Encoding.ASCII.GetString(data, record, 4), (int)offset, (int)length));
        }
        return true;
    }

    private static bool TryReadCodeSubtables(byte[] data, int cmap, int length, out int? symbolic, out int? macintosh,
        out bool hasUnicodeCmap) {
        symbolic = null;
        macintosh = null;
        hasUnicodeCmap = false;
        if (length < 4) return false;
        int count = ReadUInt16(data, cmap + 2);
        if (count > OfficeOpenTypeCmap.MaximumSubtables || 4 + count * 8 > length) return false;
        HashSet<int> validFormat4 = OfficeOpenTypeCmap.CollectValidFormat4Subtables(
            data, cmap, length, OfficeOpenTypeCmap.MaximumSubtables);
        HashSet<int> validFormat12 = OfficeOpenTypeCmap.CollectValidFormat12Subtables(
            data, cmap, length, OfficeOpenTypeCmap.MaximumSubtables, OfficeOpenTypeCmap.MaximumFormat12Groups);
        for (int index = 0; index < count; index++) {
            int record = cmap + 4 + index * 8;
            int platform = ReadUInt16(data, record);
            int encoding = ReadUInt16(data, record + 2);
            uint offset = ReadUInt32(data, record + 4);
            if (offset > (uint)(length - 4)) continue;
            int table = cmap + (int)offset;
            int format = ReadUInt16(data, table);
            if (platform == 0 || platform == 3 && (encoding == 1 || encoding == 10)) {
                hasUnicodeCmap |= format == 4 && validFormat4.Contains(table) ||
                    format == 12 && validFormat12.Contains(table);
                continue;
            }
            if (format != 0 && format != 4 && format != 6 || !HasSubtableBody(data, table, cmap + length)) continue;
            if (platform == 3 && encoding == 0) symbolic ??= table;
            else if (platform == 1 && encoding == 0) macintosh ??= table;
        }
        return true;
    }

    private static bool HasSubtableBody(byte[] data, int table, int cmapEnd) {
        if (table > cmapEnd - 4) return false;
        int length = ReadUInt16(data, table + 2);
        return length >= 6 && table <= cmapEnd - length;
    }

    private static int LookupSubtable(byte[] data, int table, int code) {
        int length = ReadUInt16(data, table + 2);
        switch (ReadUInt16(data, table)) {
            case 0:
                return code < 256 && length >= 262 ? data[table + 6 + code] : 0;
            case 6: {
                if (length < 10) return 0;
                int first = ReadUInt16(data, table + 6);
                int entries = ReadUInt16(data, table + 8);
                if (10 + entries * 2 > length || code < first || code >= first + entries) return 0;
                return ReadUInt16(data, table + 10 + (code - first) * 2);
            }
            case 4: {
                if (length < 16) return 0;
                int segments = ReadUInt16(data, table + 6) / 2;
                int ends = table + 14;
                int starts = ends + segments * 2 + 2;
                int deltas = starts + segments * 2;
                int rangeOffsets = deltas + segments * 2;
                if (rangeOffsets + segments * 2 > table + length) return 0;
                for (int segment = 0; segment < segments; segment++) {
                    if (code > ReadUInt16(data, ends + segment * 2)) continue;
                    int start = ReadUInt16(data, starts + segment * 2);
                    if (code < start) return 0;
                    int delta = ReadUInt16(data, deltas + segment * 2);
                    int rangeOffsetPosition = rangeOffsets + segment * 2;
                    int rangeOffset = ReadUInt16(data, rangeOffsetPosition);
                    if (rangeOffset == 0) return (code + delta) & 0xFFFF;
                    int glyphPosition = rangeOffsetPosition + rangeOffset + (code - start) * 2;
                    if (glyphPosition > table + length - 2) return 0;
                    int glyph = ReadUInt16(data, glyphPosition);
                    return glyph == 0 ? 0 : (glyph + delta) & 0xFFFF;
                }
                return 0;
            }
            default:
                return 0;
        }
    }

    private static bool IsWhiteSpaceScalar(int scalar) => scalar <= 0xFFFF && char.IsWhiteSpace((char)scalar);

    // A glyf glyph is empty when its loca range has no data.
    private static Func<int, bool> CreateEmptyGlyphTest(byte[] data, List<(string Tag, int Offset, int Length)> tables, int glyphCount) {
        var head = tables.Find(static table => string.Equals(table.Tag, "head", StringComparison.Ordinal));
        var loca = tables.Find(static table => string.Equals(table.Tag, "loca", StringComparison.Ordinal));
        if (head.Length < 54 || loca.Length == 0) return static _ => false;
        bool longOffsets = ReadUInt16(data, head.Offset + 50) != 0;
        int entrySize = longOffsets ? 4 : 2;
        return glyph => {
            if (glyph < 0 || glyph >= glyphCount || (glyph + 2) * entrySize > loca.Length) return false;
            int offset = loca.Offset + glyph * entrySize;
            uint start = longOffsets ? ReadUInt32(data, offset) : (uint)ReadUInt16(data, offset) * 2;
            uint end = longOffsets ? ReadUInt32(data, offset + entrySize) : (uint)ReadUInt16(data, offset + entrySize) * 2;
            return end <= start;
        };
    }

    private static int ReadGlyphCount(byte[] data, List<(string Tag, int Offset, int Length)> tables) {
        foreach (var table in tables) {
            if (string.Equals(table.Tag, "maxp", StringComparison.Ordinal) && table.Length >= 6) return ReadUInt16(data, table.Offset + 4);
        }
        return 0;
    }

    internal static byte[] BuildUnicodeCmap(SortedDictionary<int, int> mappings) {
        var basic = mappings.Where(static mapping => mapping.Key <= 0xFFFF).ToList();
        // Format 4 uses one segment per mapping plus the sentinel. The shared reader
        // accepts at most MaximumSubtables * 16 segments, below the ushort length limit.
        byte[]? format4 = basic.Count < OfficeOpenTypeCmap.MaximumSubtables * 16 ? BuildFormat4(basic) : null;
        bool needsFull = format4 == null || basic.Count != mappings.Count;
        List<byte[]> format12 = needsFull ? BuildFormat12Parts(mappings) : new List<byte[]>();
        using var output = new MemoryStream();
        int records = (format4 == null ? 0 : 1) + format12.Count;
        WriteUInt16(output, 0);
        WriteUInt16(output, (ushort)records);
        int offset = 4 + records * 8;
        if (format4 != null) {
            WriteUInt16(output, 3);
            WriteUInt16(output, 1);
            WriteUInt32(output, (uint)offset);
        }
        int fullOffset = offset + (format4?.Length ?? 0);
        foreach (byte[] table in format12) {
            WriteUInt16(output, 3);
            WriteUInt16(output, 10);
            WriteUInt32(output, (uint)fullOffset);
            fullOffset += table.Length;
        }
        if (format4 != null) output.Write(format4, 0, format4.Length);
        foreach (byte[] table in format12) output.Write(table, 0, table.Length);
        return output.ToArray();
    }

    // One segment per mapped character keeps the table simple; simple fonts map at most 256 codes.
    private static byte[] BuildFormat4(List<KeyValuePair<int, int>> mappings) {
        int segments = mappings.Count + 1;
        int searchRange = 2;
        int entrySelector = 0;
        while (searchRange * 2 <= segments * 2) {
            searchRange *= 2;
            entrySelector++;
        }
        using var output = new MemoryStream();
        WriteUInt16(output, 4);
        WriteUInt16(output, (ushort)(16 + segments * 8));
        WriteUInt16(output, 0);
        WriteUInt16(output, (ushort)(segments * 2));
        WriteUInt16(output, (ushort)searchRange);
        WriteUInt16(output, (ushort)entrySelector);
        WriteUInt16(output, (ushort)(segments * 2 - searchRange));
        foreach (var mapping in mappings) WriteUInt16(output, (ushort)mapping.Key);
        WriteUInt16(output, 0xFFFF);
        WriteUInt16(output, 0);
        foreach (var mapping in mappings) WriteUInt16(output, (ushort)mapping.Key);
        WriteUInt16(output, 0xFFFF);
        foreach (var mapping in mappings) WriteUInt16(output, unchecked((ushort)(mapping.Value - mapping.Key)));
        WriteUInt16(output, 1);
        for (int segment = 0; segment < segments; segment++) WriteUInt16(output, 0);
        return output.ToArray();
    }

    private static List<byte[]> BuildFormat12Parts(SortedDictionary<int, int> mappings) {
        var groups = new List<(int Start, int End, int Glyph)>();
        foreach (var mapping in mappings) {
            if (groups.Count > 0) {
                var previous = groups[groups.Count - 1];
                if (mapping.Key == previous.End + 1 && mapping.Value == previous.Glyph + mapping.Key - previous.Start) {
                    groups[groups.Count - 1] = (previous.Start, mapping.Key, previous.Glyph);
                    continue;
                }
            }
            groups.Add((mapping.Key, mapping.Key, mapping.Value));
        }
        int maximumGroups = (int)OfficeOpenTypeCmap.MaximumFormat12Groups;
        var parts = new List<byte[]>((groups.Count + maximumGroups - 1) / maximumGroups);
        for (int first = 0; first < groups.Count; first += maximumGroups) {
            int count = Math.Min(maximumGroups, groups.Count - first);
            using var output = new MemoryStream();
            WriteUInt16(output, 12);
            WriteUInt16(output, 0);
            WriteUInt32(output, (uint)(16 + count * 12));
            WriteUInt32(output, 0);
            WriteUInt32(output, (uint)count);
            for (int index = first; index < first + count; index++) {
                var group = groups[index];
                WriteUInt32(output, (uint)group.Start);
                WriteUInt32(output, (uint)group.End);
                WriteUInt32(output, (uint)group.Glyph);
            }
            parts.Add(output.ToArray());
        }
        return parts;
    }

    private static byte[] Assemble(List<(string Tag, byte[] Body)> tables) {
        int count = tables.Count;
        int power = 1;
        int entrySelector = 0;
        while (power * 2 <= count) {
            power *= 2;
            entrySelector++;
        }
        using var output = new MemoryStream();
        WriteUInt32(output, 0x00010000);
        WriteUInt16(output, (ushort)count);
        WriteUInt16(output, (ushort)(power * 16));
        WriteUInt16(output, (ushort)entrySelector);
        WriteUInt16(output, (ushort)(count * 16 - power * 16));
        output.Write(new byte[count * 16], 0, count * 16);
        var offsets = new int[count];
        int headOffset = -1;
        for (int index = 0; index < count; index++) {
            while (output.Length % 4 != 0) output.WriteByte(0);
            offsets[index] = (int)output.Length;
            byte[] body = tables[index].Body;
            if (string.Equals(tables[index].Tag, "head", StringComparison.Ordinal) && body.Length >= 12) {
                headOffset = offsets[index];
                WriteUInt32(body, 8, 0);
            }
            output.Write(body, 0, body.Length);
        }
        while (output.Length % 4 != 0) output.WriteByte(0);
        byte[] result = output.ToArray();
        for (int index = 0; index < count; index++) {
            int record = 12 + index * 16;
            Encoding.ASCII.GetBytes(tables[index].Tag, 0, 4, result, record);
            WriteUInt32(result, record + 4, Checksum(tables[index].Body));
            WriteUInt32(result, record + 8, (uint)offsets[index]);
            WriteUInt32(result, record + 12, (uint)tables[index].Body.Length);
        }
        if (headOffset >= 0) WriteUInt32(result, headOffset + 8, unchecked(ChecksumMagic - Checksum(result)));
        return result;
    }

    private static uint Checksum(byte[] data) {
        uint sum = 0;
        for (int offset = 0; offset < data.Length; offset += 4) {
            uint value = 0;
            for (int index = 0; index < 4; index++) {
                value = (value << 8) | (offset + index < data.Length ? data[offset + index] : 0u);
            }
            sum = unchecked(sum + value);
        }
        return sum;
    }

    private static int ReadUInt16(byte[] data, int offset) => (data[offset] << 8) | data[offset + 1];

    private static uint ReadUInt32(byte[] data, int offset) =>
        ((uint)data[offset] << 24) | ((uint)data[offset + 1] << 16) | ((uint)data[offset + 2] << 8) | data[offset + 3];

    private static void WriteUInt16(Stream output, ushort value) {
        output.WriteByte((byte)(value >> 8));
        output.WriteByte((byte)value);
    }

    private static void WriteUInt32(Stream output, uint value) {
        output.WriteByte((byte)(value >> 24));
        output.WriteByte((byte)(value >> 16));
        output.WriteByte((byte)(value >> 8));
        output.WriteByte((byte)value);
    }

    private static void WriteUInt32(byte[] data, int offset, uint value) {
        data[offset] = (byte)(value >> 24);
        data[offset + 1] = (byte)(value >> 16);
        data[offset + 2] = (byte)(value >> 8);
        data[offset + 3] = (byte)value;
    }
}

/// <summary>A drawing-ready TrueType program with its synthesized Unicode mappings.</summary>
internal sealed class PdfDrawingFontProgram {
    internal PdfDrawingFontProgram(byte[] program, SortedDictionary<int, int> unicodeGlyphs, Func<int, int> glyphForCode,
        Func<int, bool> isEmptyGlyph) {
        Program = program;
        UnicodeGlyphs = unicodeGlyphs;
        GlyphForCode = glyphForCode;
        IsEmptyGlyph = isEmptyGlyph;
    }

    internal byte[] Program { get; }

    /// <summary>Unicode scalar to glyph id entries of the synthesized cmap.</summary>
    internal IReadOnlyDictionary<int, int> UnicodeGlyphs { get; }

    /// <summary>Resolves the glyph a PDF character code or CID paints, or 0.</summary>
    internal Func<int, int> GlyphForCode { get; }

    /// <summary>True when a glyph has no outline data.</summary>
    internal Func<int, bool> IsEmptyGlyph { get; }
}
