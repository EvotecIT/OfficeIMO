using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal sealed partial class PdfOpenTypeCffFontProgram {
    private const uint OpenTypeCffScalerType = 0x4F54544F;

    private readonly byte[] _data;
    private readonly ushort[] _advanceWidths;
    private readonly Dictionary<int, int> _cmap;
    private readonly Dictionary<string, TableRecord> _tables;
    private readonly SortedSet<int> _usedGlyphIds = new();
    private readonly Dictionary<int, string> _usedGlyphToUnicode = new();
    private readonly object _usageLock = new();

    private PdfOpenTypeCffFontProgram(
        byte[] data,
        string fontName,
        int unitsPerEm,
        int xMin,
        int yMin,
        int xMax,
        int yMax,
        int ascent,
        int descent,
        int capHeight,
        double italicAngle,
        int flags,
        int stemV,
        ushort[] advanceWidths,
        Dictionary<int, int> cmap,
        Dictionary<string, TableRecord> tables,
        int cffTableLength) {
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
        CffTableLength = cffTableLength;
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
    public int GlyphCount => _advanceWidths.Length;
    public int CffTableLength { get; }
    public int FontDataLength => _data.Length;
    internal byte[] FontDataSnapshot => _data.ToArray();
    internal byte[] FontDataForInspection => _data;

    public static PdfOpenTypeCffFontProgram Parse(byte[] data, string? fontNameOverride = null) {
        Guard.NotNull(data, nameof(data));
        if (data.Length < 12) {
            throw new NotSupportedException("OpenType/CFF font data is too small to embed.");
        }

        uint scalerType = ReadUInt32(data, 0);
        if (scalerType != OpenTypeCffScalerType) {
            throw new NotSupportedException("Only OpenType/CFF fonts with an OTTO scaler can be parsed by this font program.");
        }

        PdfOpenTypeFontInfo info = PdfOpenTypeFontInspector.Inspect(data, fontNameOverride);
        if (!info.IsOpenTypeCff) {
            throw new NotSupportedException("OpenType font does not contain CFF outlines.");
        }

        Dictionary<string, TableRecord> tables = ReadTableDirectory(data);
        TableRecord head = GetTable(tables, "head");
        TableRecord hhea = GetTable(tables, "hhea");
        TableRecord hmtx = GetTable(tables, "hmtx");
        TableRecord cff = GetTable(tables, "CFF ");
        TableRecord? os2 = tables.TryGetValue("OS/2", out TableRecord os2Record) ? os2Record : null;
        TableRecord? post = tables.TryGetValue("post", out TableRecord postRecord) ? postRecord : null;

        int unitsPerEm = ReadUInt16(data, head.Offset + 18);
        if (unitsPerEm <= 0) {
            throw new NotSupportedException("OpenType/CFF font has an invalid unitsPerEm value.");
        }

        int xMin = ReadInt16(data, head.Offset + 36);
        int yMin = ReadInt16(data, head.Offset + 38);
        int xMax = ReadInt16(data, head.Offset + 40);
        int yMax = ReadInt16(data, head.Offset + 42);
        int macStyle = ReadUInt16(data, head.Offset + 44);
        int ascent = ReadInt16(data, hhea.Offset + 4);
        int descent = ReadInt16(data, hhea.Offset + 6);
        int numberOfHMetrics = ReadUInt16(data, hhea.Offset + 34);
        if (numberOfHMetrics <= 0 || info.GlyphCount <= 0) {
            throw new NotSupportedException("OpenType/CFF font horizontal metrics are invalid.");
        }

        int capHeight = ascent;
        int weightClass = 400;
        if (os2.HasValue) {
            ushort fsType = ReadUInt16(data, os2.Value.Offset + 8);
            if ((fsType & 0x0002) != 0) {
                throw new NotSupportedException("OpenType/CFF font embedding is restricted by the font fsType license flags.");
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

        ushort[] widths = ReadAdvanceWidths(data, hmtx, numberOfHMetrics, info.GlyphCount);
        int flags = 32;
        if (fixedPitch) flags |= 1;
        if ((macStyle & 0x02) != 0 || Math.Abs(italicAngle) > 0.01D) flags |= 64;
        int stemV = Math.Max(50, Math.Min(220, 80 + ((weightClass - 400) / 10)));

        return new PdfOpenTypeCffFontProgram(
            data,
            info.FontName,
            unitsPerEm,
            xMin,
            yMin,
            xMax,
            yMax,
            ascent,
            descent,
            capHeight,
            italicAngle,
            flags,
            stemV,
            widths,
            CopyUnicodeCMap(info.UnicodeCMap),
            tables,
            cff.Length);
    }

    public bool TryGetGlyphId(int unicodeScalar, out int glyphId) =>
        _cmap.TryGetValue(unicodeScalar, out glyphId);

    public int GetGlyphWidth1000(int glyphId) {
        if (glyphId < 0 || glyphId >= _advanceWidths.Length) {
            return _advanceWidths.Length == 0 ? 500 : ScaleMetric(_advanceWidths[_advanceWidths.Length - 1], UnitsPerEm);
        }

        return ScaleMetric(_advanceWidths[glyphId], UnitsPerEm);
    }

    public double MeasureTextWidth(string? text, double fontSize, PdfTextShapingMode shapingMode = PdfTextShapingMode.UnicodeScalar, IOfficeTextShapingProvider? shapingProvider = null, string? language = null) {
        if (string.IsNullOrEmpty(text)) {
            return 0D;
        }

        return ShapeText(text!, PdfTextShapingOptions.ForRendering(FontName, shapingMode, shapingProvider, language: language)).TotalAdvanceWidth1000 * fontSize / 1000D;
    }

    public string EncodeTextAsGlyphHex(string text, PdfTextShapingMode shapingMode = PdfTextShapingMode.UnicodeScalar, IOfficeTextShapingProvider? shapingProvider = null) {
        Guard.NotNull(text, nameof(text));
        return ShapeText(text, PdfTextShapingOptions.ForRendering(FontName, shapingMode, shapingProvider)).ToGlyphHex();
    }

    internal string EncodeTextAsGlyphHex(string text, PdfTextShapingMode shapingMode, IOfficeTextShapingProvider? shapingProvider, Action<string, string, bool>? providerShapedTextRecorder, string? language = null) {
        Guard.NotNull(text, nameof(text));
        return ShapeText(text, PdfTextShapingOptions.ForRendering(FontName, shapingMode, shapingProvider, providerShapedTextRecorder, language)).ToGlyphHex();
    }

    internal PdfGlyphRun ShapeText(string text, PdfTextShapingOptions options) {
        Guard.NotNull(text, nameof(text));
        if (PdfExternalTextShaper.TryShapeText(text, this, options, out PdfGlyphRun glyphRun)) {
            return glyphRun;
        }

        var glyphs = new List<PdfGlyphInfo>();
        for (int index = 0; index < text.Length;) {
            int scalarStart = index;
            if (options.ShapingMode == PdfTextShapingMode.LatinLigatures &&
                OfficeTextLigatures.TryGetLatinPresentationForm(text, scalarStart, out int ligatureScalar, out int ligatureLength) &&
                TryGetGlyphId(ligatureScalar, out int ligatureGlyphId) &&
                ligatureGlyphId > 0) {
                string unicodeText = text.Substring(scalarStart, ligatureLength);
                if (options.RecordGlyphUsage) {
                    RecordGlyphUsage(ligatureGlyphId, unicodeText);
                }

                glyphs.Add(new PdfGlyphInfo(ligatureGlyphId, unicodeText, scalarStart, GetGlyphWidth1000(ligatureGlyphId)));
                index += ligatureLength;
                continue;
            }

            int scalar = ReadScalar(text, ref index);
            if (!TryGetGlyphId(scalar, out int glyphId) || glyphId <= 0) {
                if (options.ThrowOnMissingGlyph) {
                    throw CreateUnsupportedGlyphException(text, scalarStart, scalar);
                }

                continue;
            }

            if (options.RecordGlyphUsage) {
                RecordGlyphUsage(glyphId, scalar);
            }

            glyphs.Add(new PdfGlyphInfo(glyphId, scalar, scalarStart, GetGlyphWidth1000(glyphId)));
        }

        return new PdfGlyphRun(glyphs);
    }

    public double GetAscender(double fontSize) =>
        Ascent * fontSize / 1000D;

    public double GetDescender(double fontSize) =>
        Math.Abs(Descent) * fontSize / 1000D;

    public byte[] BuildFullOpenTypeFontFile() =>
        _data.ToArray();

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
        foreach (KeyValuePair<int, int> entry in _cmap) {
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
        lock (_usageLock) {
            _usedGlyphIds.Add(glyphId);
            if (!string.IsNullOrEmpty(unicodeText) &&
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

    private static Dictionary<string, TableRecord> ReadTableDirectory(byte[] data) {
        int numTables = ReadUInt16(data, 4);
        var tables = new Dictionary<string, TableRecord>(StringComparer.Ordinal);
        int recordOffset = 12;
        for (int index = 0; index < numTables; index++) {
            int offset = recordOffset + index * 16;
            EnsureRange(data, offset, 16);
            string tag = Encoding.ASCII.GetString(data, offset, 4);
            uint tableOffset = ReadUInt32(data, offset + 8);
            uint tableLength = ReadUInt32(data, offset + 12);
            if (tableOffset > int.MaxValue || tableLength > int.MaxValue) {
                throw new NotSupportedException("OpenType/CFF font table offsets are too large.");
            }

            EnsureRange(data, (int)tableOffset, (int)tableLength);
            tables[tag] = new TableRecord((int)tableOffset, (int)tableLength);
        }

        return tables;
    }

    private static Dictionary<int, int> CopyUnicodeCMap(IReadOnlyDictionary<int, int> unicodeCMap) {
        var copy = new Dictionary<int, int>(unicodeCMap.Count);
        foreach (KeyValuePair<int, int> entry in unicodeCMap) {
            copy[entry.Key] = entry.Value;
        }

        return copy;
    }

    private static TableRecord GetTable(Dictionary<string, TableRecord> tables, string tag) {
        if (!tables.TryGetValue(tag, out TableRecord record)) {
            throw new NotSupportedException("OpenType/CFF font is missing required '" + tag + "' table.");
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

    private static int ScaleMetric(int value, int unitsPerEm) =>
        (int)Math.Round(value * 1000D / unitsPerEm, MidpointRounding.AwayFromZero);

    private static int ReadScalar(string text, ref int index) {
        char ch = text[index++];
        if (char.IsHighSurrogate(ch)) {
            if (index < text.Length && char.IsLowSurrogate(text[index])) {
                return char.ConvertToUtf32(ch, text[index++]);
            }

            throw new ArgumentException("Text contains an unmatched high surrogate at index " + (index - 1).ToString(CultureInfo.InvariantCulture) + ".", nameof(text));
        }

        if (char.IsLowSurrogate(ch)) {
            throw new ArgumentException("Text contains an unmatched low surrogate at index " + (index - 1).ToString(CultureInfo.InvariantCulture) + ".", nameof(text));
        }

        return ch;
    }

    private static ArgumentException CreateUnsupportedGlyphException(string text, int index, int scalar) {
        string codePoint = "U+" + scalar.ToString("X", CultureInfo.InvariantCulture);
        string display = scalar <= 0x10FFFF ? char.ConvertFromUtf32(scalar) : string.Empty;
        string rendered = display.Length == 0 || char.IsControl(display, 0) ? string.Empty : " '" + display + "'";
        return new ArgumentException("Text contains character " + codePoint + rendered + " at index " + index.ToString(CultureInfo.InvariantCulture) + " that is not covered by the embedded OpenType/CFF font.", nameof(text));
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
            throw new NotSupportedException("OpenType/CFF font table data is truncated or invalid.");
        }
    }

    private readonly struct TableRecord {
        public TableRecord(int offset, int length) {
            Offset = offset;
            Length = length;
        }

        public int Offset { get; }
        public int Length { get; }
    }
}
