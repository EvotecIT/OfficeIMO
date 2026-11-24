namespace OfficeIMO.Pdf;

internal static class PdfExtractedFontProgram {
    internal static string Extension(string? subtype, byte[] bytes) => subtype switch {
        "TrueType" or "OpenType" when HasTag(bytes, "ttcf") => "ttc",
        "TrueType" or "OpenType" when HasTag(bytes, "OTTO") => "otf",
        "TrueType" or "OpenType" when HasTag(bytes, "true") ||
            bytes.Length >= 4 && bytes[0] == 0 && bytes[1] == 1 && bytes[2] == 0 && bytes[3] == 0 => "ttf",
        "Type1C" or "CIDFontType0C" => "cff",
        "Type1" when bytes.Length >= 2 && bytes[0] == 0x80 && bytes[1] == 1 => "pfb",
        "Type1" when bytes.All(static value => value < 128) => "pfa",
        "Type1" => "type1",
        _ => "bin"
    };

    internal static byte[] RepairName(byte[] bytes, PdfFontInfo font, out bool repaired, out bool readable) {
        repaired = false;
        readable = false;
        if (!PdfTrueTypeUnicodeCmap.TryReadTables(bytes, out var tables, ignoreOutOfRangeNameTable: true) ||
            tables.Select(static table => table.Tag).Distinct(StringComparer.Ordinal).Count() != tables.Count ||
            !tables.Any(static table => table.Tag == "head" && table.Length >= 54)) return bytes;

        readable = true;
        foreach ((string tag, int offset, int length) in tables) {
            if (tag == "name" && HasUsableName(bytes, offset, length)) return bytes;
        }

        byte[] name = BuildName(font);
        var segments = new List<(string Tag, byte[] Source, int Offset, int Length)>(tables.Count + 1);
        foreach ((string tag, int offset, int length) in tables) {
            // Rebuilding the directory invalidates any digital signature on the original font.
            if (tag is not ("name" or "DSIG")) segments.Add((tag, bytes, offset, length));
        }
        segments.Add(("name", name, 0, name.Length));
        segments.Sort(static (a, b) => string.CompareOrdinal(a.Tag, b.Tag));
        uint flavor = ((uint)bytes[0] << 24) | ((uint)bytes[1] << 16) | ((uint)bytes[2] << 8) | bytes[3];
        repaired = true;
        return PdfTrueTypeUnicodeCmap.Assemble(segments, flavor);
    }

    private static bool HasUsableName(byte[] bytes, int start, int length) {
        if (length < 6 || Read16(bytes, start) > 1) return false;
        int count = Read16(bytes, start + 2);
        int strings = Read16(bytes, start + 4);
        int recordEnd = 6 + count * 12;
        if (strings < recordEnd || strings > length) return false;
        if (Read16(bytes, start) == 1 &&
            (recordEnd + 2 > strings || recordEnd + 2 + 4 * Read16(bytes, start + recordEnd) > strings)) return false;

        var usable = new Dictionary<(int Platform, int Encoding, int Language), int>();
        const int required = (1 << 1) | (1 << 2) | (1 << 4) | (1 << 6);
        for (int i = 0; i < count; i++) {
            int record = start + 6 + i * 12;
            int platform = Read16(bytes, record);
            int encoding = Read16(bytes, record + 2);
            int language = Read16(bytes, record + 4);
            int id = Read16(bytes, record + 6);
            int size = Read16(bytes, record + 8);
            int offset = Read16(bytes, record + 10);
            if (offset + size > length - strings) return false;
            if (id is not (1 or 2 or 4 or 6) || size == 0) continue;
            int position = start + strings + offset;
            string text;
            if (platform == 1 && encoding == 0) text = PdfEncoding.Latin1GetString(bytes.AsSpan(position, size));
            else if ((platform == 0 || platform == 3 && encoding is 0 or 1 or 10) && size % 2 == 0) {
                text = Encoding.BigEndianUnicode.GetString(bytes, position, size);
            } else continue;
            if (string.IsNullOrWhiteSpace(text) || text.Any(static c => char.IsControl(c) || c == '\uFFFD')) continue;
            if (id == 6 && (text.Length > 63 || text.Any(static c => c is < '!' or > '~' or '[' or ']' or '(' or ')' or '{' or '}' or '<' or '>' or '/' or '%'))) continue;
            var key = (platform, encoding, language);
            usable.TryGetValue(key, out int mask);
            usable[key] = mask | 1 << id;
        }
        return usable.Values.Any(static mask => (mask & required) == required);
    }

    private static byte[] BuildName(PdfFontInfo font) {
        string family = new string(font.FamilyName.Where(static c => !char.IsControl(c)).Take(96).ToArray()).Trim();
        if (family.Length == 0) family = "Embedded PDF Font";
        if (font.SubsetTag is not null) family += " " + font.SubsetTag;
        string style = font.IsBold ? (font.IsItalic ? "Bold Italic" : "Bold") : (font.IsItalic ? "Italic" : "Regular");
        string fullName = style == "Regular" ? family : family + " " + style;
        string postScriptFamily = new string(family.Where(static c => c is >= 'A' and <= 'Z' or >= 'a' and <= 'z' or >= '0' and <= '9').Take(48).ToArray());
        if (postScriptFamily.Length == 0) postScriptFamily = "EmbeddedPdfFont";
        string postScriptName = postScriptFamily + "-" + style.Replace(" ", string.Empty);
        (int Id, string Text)[] names = {
            (1, family), (2, style), (3, postScriptName + ";1.0"), (4, fullName), (5, "Version 1.0"), (6, postScriptName)
        };

        int strings = 6 + names.Length * 12;
        byte[] table = new byte[strings + names.Sum(static name => Encoding.BigEndianUnicode.GetByteCount(name.Text))];
        Write16(table, 2, names.Length);
        Write16(table, 4, strings);
        int offset = 0;
        for (int i = 0; i < names.Length; i++) {
            int record = 6 + i * 12;
            int size = Encoding.BigEndianUnicode.GetBytes(names[i].Text, 0, names[i].Text.Length, table, strings + offset);
            Write16(table, record, 3); // Windows
            Write16(table, record + 2, 1); // Unicode BMP
            Write16(table, record + 4, 0x0409); // English
            Write16(table, record + 6, names[i].Id);
            Write16(table, record + 8, size);
            Write16(table, record + 10, offset);
            offset += size;
        }
        return table;
    }

    private static bool HasTag(byte[] bytes, string tag) =>
        bytes.Length >= 4 && bytes[0] == tag[0] && bytes[1] == tag[1] && bytes[2] == tag[2] && bytes[3] == tag[3];

    private static int Read16(byte[] bytes, int offset) => bytes[offset] << 8 | bytes[offset + 1];

    private static void Write16(byte[] bytes, int offset, int value) {
        bytes[offset] = (byte)(value >> 8);
        bytes[offset + 1] = (byte)value;
    }
}
