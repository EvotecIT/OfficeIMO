using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal sealed partial class PdfTrueTypeFontProgram {
    private const int MaxUnicodeCMapMappings = 131_072;

    private readonly byte[] _data;
    private readonly ushort[] _advanceWidths;
    private readonly Dictionary<int, int> _cmap;
    private readonly Dictionary<string, TableRecord> _tables;
    private readonly SortedSet<int> _usedGlyphIds = new();
    private readonly Dictionary<int, string> _usedGlyphToUnicode = new();
    private readonly object _usageLock = new();

    private PdfTrueTypeFontProgram(byte[] data, Dictionary<string, TableRecord> tables, string fontName, int unitsPerEm, int xMin, int yMin, int xMax, int yMax, int ascent, int descent, int capHeight, double italicAngle, int flags, int stemV, ushort[] advanceWidths, Dictionary<int, int> cmap) {
        _data = data.ToArray();
        _tables = new Dictionary<string, TableRecord>(tables, StringComparer.Ordinal);
        FontName = fontName;
        UnitsPerEm = unitsPerEm;
        FontBBox = new[] { ScaleMetric(xMin, unitsPerEm), ScaleMetric(yMin, unitsPerEm), ScaleMetric(xMax, unitsPerEm), ScaleMetric(yMax, unitsPerEm) };
        Ascent = ScaleMetric(ascent, unitsPerEm);
        Descent = ScaleMetric(descent, unitsPerEm);
        CapHeight = ScaleMetric(capHeight, unitsPerEm);
        ItalicAngle = italicAngle;
        Flags = flags;
        StemV = stemV;
        _advanceWidths = advanceWidths;
        _cmap = cmap;
    }

    public string FontName { get; }
    public int UnitsPerEm { get; }
    public int[] FontBBox { get; }
    public int Ascent { get; }
    public int Descent { get; }
    public int CapHeight { get; }
    public double ItalicAngle { get; }
    public int Flags { get; }
    public int StemV { get; }

    public double MeasureWinAnsiTextWidth(string? text, double fontSize) {
        if (string.IsNullOrEmpty(text)) {
            return 0D;
        }

        double width = 0D;
        for (int index = 0; index < text!.Length; index++) {
            width += GetWinAnsiGlyphWidth1000(text[index]) * fontSize / 1000D;
        }

        return width;
    }

    public double MeasureTextWidth(string? text, double fontSize, PdfTextShapingMode shapingMode = PdfTextShapingMode.UnicodeScalar, IOfficeTextShapingProvider? shapingProvider = null, string? language = null) {
        if (string.IsNullOrEmpty(text)) {
            return 0D;
        }

        return ShapeText(text!, PdfTextShapingOptions.ForRendering(FontName, shapingMode, shapingProvider, language: language)).TotalAdvanceWidth1000 * fontSize / 1000D;
    }

    public double GetAscender(double fontSize) =>
        Ascent * fontSize / 1000D;

    public double GetDescender(double fontSize) =>
        Math.Abs(Descent) * fontSize / 1000D;

    public int GlyphCount => _advanceWidths.Length;

    internal byte[] FontDataForInspection => _data;

    public bool TryGetGlyphId(int unicodeScalar, out int glyphId) =>
        _cmap.TryGetValue(unicodeScalar, out glyphId);

    public int GetGlyphWidth1000(int glyphId) {
        if (glyphId < 0 || glyphId >= _advanceWidths.Length) {
            return _advanceWidths.Length == 0 ? 500 : ScaleMetric(_advanceWidths[_advanceWidths.Length - 1], UnitsPerEm);
        }

        return ScaleMetric(_advanceWidths[glyphId], UnitsPerEm);
    }

    public string EncodeTextAsGlyphHex(string text, PdfTextShapingMode shapingMode = PdfTextShapingMode.UnicodeScalar, IOfficeTextShapingProvider? shapingProvider = null) {
        Guard.NotNull(text, nameof(text));
        return ShapeText(text, PdfTextShapingOptions.ForRendering(FontName, shapingMode, shapingProvider)).ToGlyphHex();
    }

    internal string EncodeTextAsGlyphHex(string text, PdfTextShapingMode shapingMode, IOfficeTextShapingProvider? shapingProvider, Action<string, string, bool>? providerShapedTextRecorder, string? language = null) {
        Guard.NotNull(text, nameof(text));
        return ShapeText(text, PdfTextShapingOptions.ForRendering(FontName, shapingMode, shapingProvider, providerShapedTextRecorder, language)).ToGlyphHex();
    }

    internal PdfGlyphRun ShapeText(string text) {
        Guard.NotNull(text, nameof(text));
        return ShapeText(text, PdfTextShapingOptions.ForRendering(FontName));
    }

    internal PdfGlyphRun ShapeText(string text, PdfTextShapingMode shapingMode) {
        Guard.NotNull(text, nameof(text));
        return ShapeText(text, PdfTextShapingOptions.ForRendering(FontName, shapingMode));
    }

    internal PdfGlyphRun ShapeText(string text, PdfTextShapingOptions options) {
        Guard.NotNull(text, nameof(text));
        if (PdfExternalTextShaper.TryShapeText(text, this, options, out PdfGlyphRun glyphRun)) {
            return glyphRun;
        }

        return PdfUnicodeScalarTextShaper.Instance.ShapeText(text, this, options);
    }

    public IReadOnlyList<(int GlyphId, string UnicodeText)> GetGlyphToUnicodeMappings() {
        lock (_usageLock) {
            if (_usedGlyphToUnicode.Count > 0) {
                return _usedGlyphToUnicode
                    .OrderBy(entry => entry.Key)
                    .Select(entry => (entry.Key, entry.Value))
                    .ToArray();
            }
        }

        var glyphToUnicode = new Dictionary<int, string>();
        var glyphToScalar = new Dictionary<int, int>();
        foreach (var entry in _cmap) {
            if (entry.Value <= 0) {
                continue;
            }

            if (!glyphToScalar.TryGetValue(entry.Value, out int existingScalar) || entry.Key < existingScalar) {
                glyphToScalar[entry.Value] = entry.Key;
                glyphToUnicode[entry.Value] = char.ConvertFromUtf32(entry.Key);
            }
        }

        return glyphToUnicode
            .OrderBy(entry => entry.Key)
            .Select(entry => (entry.Key, entry.Value))
            .ToArray();
    }

    internal IReadOnlyList<int> GetUsedGlyphIds() {
        lock (_usageLock) {
            return _usedGlyphIds.Count == 0
                ? Array.Empty<int>()
                : _usedGlyphIds.ToArray();
        }
    }

    internal void ResetGlyphUsage() {
        lock (_usageLock) {
            _usedGlyphIds.Clear();
            _usedGlyphToUnicode.Clear();
        }
    }

    internal void RecordGlyphUsage(int glyphId, int unicodeScalar) =>
        RecordGlyphUsage(glyphId, char.ConvertFromUtf32(unicodeScalar));

    internal void RecordGlyphUsage(int glyphId, string unicodeText) {
        if (glyphId < 0) {
            return;
        }

        lock (_usageLock) {
            _usedGlyphIds.Add(glyphId);
            if (glyphId > 0 &&
                !string.IsNullOrEmpty(unicodeText) &&
                (!_usedGlyphToUnicode.TryGetValue(glyphId, out string? existingText) || ShouldReplaceGlyphUnicodeText(unicodeText, existingText))) {
                _usedGlyphToUnicode[glyphId] = unicodeText;
            }
        }
    }

    private static bool ShouldReplaceGlyphUnicodeText(string candidate, string existing) {
        if (candidate.Length != existing.Length) {
            return candidate.Length > existing.Length;
        }

        return string.CompareOrdinal(candidate, existing) < 0;
    }

    public static PdfTrueTypeFontProgram Parse(byte[] data, string? fontNameOverride = null) {
        Guard.NotNull(data, nameof(data));
        if (data.Length < 12) {
            throw new NotSupportedException("TrueType font data is too small to embed.");
        }

        uint scalerType = ReadUInt32(data, 0);
        if (scalerType == 0x4F54544F) {
            throw new NotSupportedException("This font program parser handles TrueType fonts with glyf outlines. Use the OpenType/CFF parser path for fonts with an OTTO scaler.");
        }

        if (scalerType != 0x00010000 && scalerType != 0x74727565) {
            throw new NotSupportedException("Only TrueType fonts with glyf outlines can be embedded by OfficeIMO.Pdf at this stage.");
        }

        var tables = ReadTableDirectory(data);
        var head = GetTable(tables, "head");
        var hhea = GetTable(tables, "hhea");
        var hmtx = GetTable(tables, "hmtx");
        var maxp = GetTable(tables, "maxp");
        var cmap = GetTable(tables, "cmap");
        var name = GetTable(tables, "name");
        TableRecord? os2 = tables.TryGetValue("OS/2", out TableRecord os2Record) ? os2Record : null;
        TableRecord? post = tables.TryGetValue("post", out TableRecord postRecord) ? postRecord : null;

        int unitsPerEm = ReadUInt16(data, head.Offset + 18);
        if (unitsPerEm <= 0) {
            throw new NotSupportedException("TrueType font has an invalid unitsPerEm value.");
        }

        int xMin = ReadInt16(data, head.Offset + 36);
        int yMin = ReadInt16(data, head.Offset + 38);
        int xMax = ReadInt16(data, head.Offset + 40);
        int yMax = ReadInt16(data, head.Offset + 42);
        int macStyle = ReadUInt16(data, head.Offset + 44);
        int ascent = ReadInt16(data, hhea.Offset + 4);
        int descent = ReadInt16(data, hhea.Offset + 6);
        int numberOfHMetrics = ReadUInt16(data, hhea.Offset + 34);
        int glyphCount = ReadUInt16(data, maxp.Offset + 4);
        if (numberOfHMetrics <= 0 || glyphCount <= 0) {
            throw new NotSupportedException("TrueType font horizontal metrics are invalid.");
        }

        int capHeight = ascent;
        int weightClass = 400;
        if (os2.HasValue) {
            ushort fsType = ReadUInt16(data, os2.Value.Offset + 8);
            if ((fsType & 0x0002) != 0) {
                throw new NotSupportedException("TrueType font embedding is restricted by the font fsType license flags.");
            }

            weightClass = ReadUInt16(data, os2.Value.Offset + 4);
            if (os2.Value.Length >= 90 && ReadUInt16(data, os2.Value.Offset) >= 2) {
                capHeight = ReadInt16(data, os2.Value.Offset + 88);
            }
        }

        double italicAngle = 0D;
        bool fixedPitch = false;
        if (post.HasValue && post.Value.Length >= 16) {
            italicAngle = ReadFixed16Dot16(data, post.Value.Offset + 4);
            fixedPitch = ReadUInt32(data, post.Value.Offset + 12) != 0;
        }

        ushort[] widths = ReadAdvanceWidths(data, hmtx, numberOfHMetrics, glyphCount);
        Dictionary<int, int> charMap = ReadUnicodeCMap(data, cmap);
        string fontName = SanitizePdfName(string.IsNullOrWhiteSpace(fontNameOverride) ? ReadPostScriptName(data, name) : fontNameOverride!);
        int flags = 32;
        if (fixedPitch) flags |= 1;
        if ((macStyle & 0x02) != 0 || Math.Abs(italicAngle) > 0.01D) flags |= 64;
        int stemV = Math.Max(50, Math.Min(220, 80 + ((weightClass - 400) / 10)));

        return new PdfTrueTypeFontProgram(data, tables, fontName, unitsPerEm, xMin, yMin, xMax, yMax, ascent, descent, capHeight, italicAngle, flags, stemV, widths, charMap);
    }

    public int[] BuildWinAnsiWidths() {
        var widths = new int[224];
        for (int code = 32; code <= 255; code++) {
            widths[code - 32] = GetWinAnsiGlyphWidth1000(PdfWinAnsiEncoding.Decode((byte)code));
        }

        return widths;
    }

    private int GetWinAnsiGlyphWidth1000(char character) {
        if (!_cmap.TryGetValue(character, out int glyphId)) {
            return 500;
        }

        ushort advance = glyphId >= 0 && glyphId < _advanceWidths.Length ? _advanceWidths[glyphId] : _advanceWidths[_advanceWidths.Length - 1];
        return ScaleMetric(advance, UnitsPerEm);
    }

    private static Dictionary<string, TableRecord> ReadTableDirectory(byte[] data) {
        int numTables = ReadUInt16(data, 4);
        var tables = new Dictionary<string, TableRecord>(StringComparer.Ordinal);
        int recordOffset = 12;
        for (int index = 0; index < numTables; index++) {
            int offset = recordOffset + index * 16;
            EnsureRange(data, offset, 16);
            string tag = Encoding.ASCII.GetString(data, offset, 4);
            uint checksum = ReadUInt32(data, offset + 4);
            uint tableOffset = ReadUInt32(data, offset + 8);
            uint tableLength = ReadUInt32(data, offset + 12);
            if (tableOffset > int.MaxValue || tableLength > int.MaxValue) {
                throw new NotSupportedException("TrueType font table offsets are too large.");
            }

            EnsureRange(data, (int)tableOffset, (int)tableLength);
            tables[tag] = new TableRecord((int)tableOffset, (int)tableLength, checksum);
        }

        return tables;
    }

    private static TableRecord GetTable(Dictionary<string, TableRecord> tables, string tag) {
        if (!tables.TryGetValue(tag, out TableRecord record)) {
            throw new NotSupportedException("TrueType font is missing required '" + tag + "' table.");
        }

        return record;
    }

    private static ushort[] ReadAdvanceWidths(byte[] data, TableRecord hmtx, int numberOfHMetrics, int glyphCount) {
        var widths = new ushort[glyphCount];
        ushort lastAdvance = 500;
        for (int glyph = 0; glyph < glyphCount; glyph++) {
            if (glyph < numberOfHMetrics) {
                int metricOffset = hmtx.Offset + glyph * 4;
                EnsureRange(data, metricOffset, 4);
                lastAdvance = ReadUInt16(data, metricOffset);
            }

            widths[glyph] = lastAdvance;
        }

        return widths;
    }

    private static Dictionary<int, int> ReadUnicodeCMap(byte[] data, TableRecord cmapTable) {
        int tableStart = cmapTable.Offset;
        int numTables = ReadUInt16(data, tableStart + 2);
        int selectedFormat12Offset = -1;
        int selectedOffset = -1;
        int fallbackOffset = -1;
        for (int i = 0; i < numTables; i++) {
            int record = tableStart + 4 + i * 8;
            EnsureRange(data, record, 8);
            int platformId = ReadUInt16(data, record);
            int encodingId = ReadUInt16(data, record + 2);
            int subtableOffset = checked(tableStart + (int)ReadUInt32(data, record + 4));
            EnsureRange(data, subtableOffset, 2);
            int format = ReadUInt16(data, subtableOffset);
            if (format == 12 && ((platformId == 3 && encodingId == 10) || platformId == 0)) {
                selectedFormat12Offset = subtableOffset;
                continue;
            }

            if (format == 4 && platformId == 3 && (encodingId == 1 || encodingId == 0) && selectedOffset < 0) {
                selectedOffset = subtableOffset;
                continue;
            }

            if (format == 4 && fallbackOffset < 0) {
                fallbackOffset = subtableOffset;
            } else if (format == 0 && fallbackOffset < 0) {
                fallbackOffset = subtableOffset;
            }
        }

        int offset = selectedFormat12Offset >= 0 ? selectedFormat12Offset : selectedOffset >= 0 ? selectedOffset : fallbackOffset;
        if (offset < 0) {
            throw new NotSupportedException("TrueType font does not contain a supported Unicode cmap subtable.");
        }

        int selectedFormat = ReadUInt16(data, offset);
        return selectedFormat == 12 ? ReadFormat12CMap(data, offset) : selectedFormat == 4 ? ReadFormat4CMap(data, offset) : ReadFormat0CMap(data, offset);
    }

    private static Dictionary<int, int> ReadFormat0CMap(byte[] data, int offset) {
        EnsureRange(data, offset, 262);
        var map = new Dictionary<int, int>();
        for (int code = 0; code < 256; code++) {
            map[code] = data[offset + 6 + code];
        }

        return map;
    }

    private static Dictionary<int, int> ReadFormat4CMap(byte[] data, int offset) {
        int length = ReadUInt16(data, offset + 2);
        EnsureRange(data, offset, length);
        int segCount = ReadUInt16(data, offset + 6) / 2;
        int endCodeOffset = offset + 14;
        int startCodeOffset = endCodeOffset + segCount * 2 + 2;
        int idDeltaOffset = startCodeOffset + segCount * 2;
        int idRangeOffsetOffset = idDeltaOffset + segCount * 2;
        var map = new Dictionary<int, int>();

        for (int segment = 0; segment < segCount; segment++) {
            int endCode = ReadUInt16(data, endCodeOffset + segment * 2);
            int startCode = ReadUInt16(data, startCodeOffset + segment * 2);
            int idDelta = ReadInt16(data, idDeltaOffset + segment * 2);
            int idRangeOffsetAddress = idRangeOffsetOffset + segment * 2;
            int idRangeOffset = ReadUInt16(data, idRangeOffsetAddress);
            if (startCode == 0xFFFF && endCode == 0xFFFF) {
                continue;
            }

            for (int code = startCode; code <= endCode && code <= 0xFFFF; code++) {
                int glyphId;
                if (idRangeOffset == 0) {
                    glyphId = (code + idDelta) & 0xFFFF;
                } else {
                    int glyphAddress = idRangeOffsetAddress + idRangeOffset + (code - startCode) * 2;
                    if (glyphAddress + 2 > offset + length) {
                        continue;
                    }

                    glyphId = ReadUInt16(data, glyphAddress);
                    if (glyphId != 0) {
                        glyphId = (glyphId + idDelta) & 0xFFFF;
                    }
                }

                if (glyphId != 0) {
                    map[code] = glyphId;
                }
            }
        }

        return map;
    }

    private static Dictionary<int, int> ReadFormat12CMap(byte[] data, int offset) {
        EnsureRange(data, offset, 16);
        uint length = ReadUInt32(data, offset + 4);
        if (length > int.MaxValue) {
            throw new NotSupportedException("TrueType format 12 cmap table is too large.");
        }

        EnsureRange(data, offset, (int)length);
        uint groupCount = ReadUInt32(data, offset + 12);
        if (groupCount > (uint)((length - 16) / 12)) {
            throw new NotSupportedException("TrueType format 12 cmap group count is invalid.");
        }

        var map = new Dictionary<int, int>();
        int groupOffset = offset + 16;
        for (uint group = 0; group < groupCount; group++) {
            uint startCharCode = ReadUInt32(data, groupOffset);
            uint endCharCode = ReadUInt32(data, groupOffset + 4);
            uint startGlyphId = ReadUInt32(data, groupOffset + 8);
            groupOffset += 12;
            if (startCharCode > 0x10FFFF || endCharCode > 0x10FFFF || endCharCode < startCharCode) {
                continue;
            }

            uint mappingCount = endCharCode - startCharCode + 1U;
            if (mappingCount > MaxUnicodeCMapMappings || map.Count > MaxUnicodeCMapMappings - (int)mappingCount) {
                throw new NotSupportedException("TrueType Unicode cmap mapping count exceeds supported limits.");
            }

            for (uint code = startCharCode; code <= endCharCode; code++) {
                uint glyph = startGlyphId + (code - startCharCode);
                if (glyph <= int.MaxValue) {
                    map[(int)code] = (int)glyph;
                }
            }
        }

        return map;
    }

    private static string ReadPostScriptName(byte[] data, TableRecord nameTable) {
        int offset = nameTable.Offset;
        int count = ReadUInt16(data, offset + 2);
        int stringOffset = offset + ReadUInt16(data, offset + 4);
        string? fallback = null;
        for (int i = 0; i < count; i++) {
            int record = offset + 6 + i * 12;
            EnsureRange(data, record, 12);
            int platformId = ReadUInt16(data, record);
            int encodingId = ReadUInt16(data, record + 2);
            int nameId = ReadUInt16(data, record + 6);
            int length = ReadUInt16(data, record + 8);
            int valueOffset = stringOffset + ReadUInt16(data, record + 10);
            EnsureRange(data, valueOffset, length);
            if (nameId != 6 && nameId != 4) {
                continue;
            }

            string value;
            if (platformId == 3 || platformId == 0) {
                value = Encoding.BigEndianUnicode.GetString(data, valueOffset, length);
            } else if (platformId == 1 && encodingId == 0) {
                value = Encoding.ASCII.GetString(data, valueOffset, length);
            } else {
                continue;
            }

            if (nameId == 6 && !string.IsNullOrWhiteSpace(value)) {
                return value;
            }

            if (fallback == null && !string.IsNullOrWhiteSpace(value)) {
                fallback = value;
            }
        }

        return fallback ?? "OfficeIMOEmbeddedFont";
    }

    private static string SanitizePdfName(string name) {
        var sb = new StringBuilder(name.Length);
        foreach (char ch in name) {
            if (char.IsLetterOrDigit(ch) || ch == '-' || ch == '_' || ch == '+') {
                sb.Append(ch);
            }
        }

        return sb.Length == 0 ? "OfficeIMOEmbeddedFont" : sb.ToString();
    }

    private static int ScaleMetric(int value, int unitsPerEm) =>
        (int)Math.Round(value * 1000D / unitsPerEm, MidpointRounding.AwayFromZero);

    private static int ReadScalar(string text, ref int index) {
        char ch = text[index++];
        if (char.IsHighSurrogate(ch)) {
            if (index < text.Length && char.IsLowSurrogate(text[index])) {
                return char.ConvertToUtf32(ch, text[index++]);
            }

            throw new ArgumentException("Text contains an unmatched high surrogate at index " + (index - 1).ToString(System.Globalization.CultureInfo.InvariantCulture) + ".", nameof(text));
        }

        if (char.IsLowSurrogate(ch)) {
            throw new ArgumentException("Text contains an unmatched low surrogate at index " + (index - 1).ToString(System.Globalization.CultureInfo.InvariantCulture) + ".", nameof(text));
        }

        return ch;
    }

    internal static ArgumentException CreateUnsupportedGlyphException(string text, int index, int scalar) {
        string codePoint = "U+" + scalar.ToString("X", System.Globalization.CultureInfo.InvariantCulture);
        string display = scalar <= 0x10FFFF ? char.ConvertFromUtf32(scalar) : string.Empty;
        string rendered = display.Length == 0 || char.IsControl(display, 0) ? string.Empty : " '" + display + "'";
        return new ArgumentException("Text contains character " + codePoint + rendered + " at index " + index.ToString(System.Globalization.CultureInfo.InvariantCulture) + " that is not covered by the embedded TrueType font.", nameof(text));
    }

    private static double ReadFixed16Dot16(byte[] data, int offset) {
        int raw = (int)ReadUInt32(data, offset);
        return raw / 65536D;
    }

    private static ushort ReadUInt16(byte[] data, int offset) {
        EnsureRange(data, offset, 2);
        return (ushort)((data[offset] << 8) | data[offset + 1]);
    }

    private static short ReadInt16(byte[] data, int offset) => unchecked((short)ReadUInt16(data, offset));

    private static uint ReadUInt32(byte[] data, int offset) {
        EnsureRange(data, offset, 4);
        return ((uint)data[offset] << 24) |
            ((uint)data[offset + 1] << 16) |
            ((uint)data[offset + 2] << 8) |
            data[offset + 3];
    }

    private static void EnsureRange(byte[] data, int offset, int length) {
        if (offset < 0 || length < 0 || offset > data.Length - length) {
            throw new NotSupportedException("TrueType font table data is truncated or invalid.");
        }
    }

    private readonly struct TableRecord {
        public TableRecord(int offset, int length, uint checksum = 0) {
            Offset = offset;
            Length = length;
            Checksum = checksum;
        }

        public int Offset { get; }
        public int Length { get; }
        public uint Checksum { get; }
    }
}
