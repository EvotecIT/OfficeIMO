using System.Globalization;

namespace OfficeIMO.Pdf;

/// <summary>
/// Provides dependency-free OpenType font inspection for PDF typography preflight.
/// </summary>
public static class PdfOpenTypeFontInspector {
    private const uint TrueTypeScalerType = 0x00010000;
    private const uint AppleTrueTypeScalerType = 0x74727565;
    private const uint OpenTypeCffScalerType = 0x4F54544F;
    private const int MaxUnicodeCMapMappings = 131_072;

    /// <summary>
    /// Parses OpenType table metadata and Unicode coverage from a single font program.
    /// </summary>
    /// <param name="fontData">OpenType or TrueType font bytes.</param>
    /// <param name="fontName">Optional configured name used when the font does not contain a readable name table.</param>
    /// <returns>Parsed font metadata and Unicode coverage.</returns>
    public static PdfOpenTypeFontInfo Inspect(byte[] fontData, string? fontName = null) {
        Guard.NotNull(fontData, nameof(fontData));
        if (fontData.Length < 12) {
            throw new NotSupportedException("OpenType font data is too small to inspect.");
        }

        uint scalerType = ReadUInt32(fontData, 0);
        if (scalerType != OpenTypeCffScalerType && scalerType != TrueTypeScalerType && scalerType != AppleTrueTypeScalerType) {
            throw new NotSupportedException("OpenType font uses unsupported scaler type " + FormatScalerType(scalerType) + ".");
        }

        Dictionary<string, TableRecord> tables = ReadTableDirectory(fontData);
        TableRecord head = GetTable(tables, "head");
        TableRecord maxp = GetTable(tables, "maxp");
        TableRecord cmap = GetTable(tables, "cmap");
        TableRecord name = GetTable(tables, "name");
        bool isOpenTypeCff = scalerType == OpenTypeCffScalerType;
        bool isTrueType = scalerType == TrueTypeScalerType || scalerType == AppleTrueTypeScalerType;
        if (isOpenTypeCff && !tables.ContainsKey("CFF ")) {
            throw new NotSupportedException("OpenType/CFF font is missing the required CFF table.");
        }

        if (isTrueType && !tables.ContainsKey("glyf")) {
            throw new NotSupportedException("TrueType font is missing the required glyf table.");
        }

        int unitsPerEm = ReadUInt16(fontData, head.Offset + 18);
        if (unitsPerEm <= 0) {
            throw new NotSupportedException("OpenType font has an invalid unitsPerEm value.");
        }

        int glyphCount = ReadUInt16(fontData, maxp.Offset + 4);
        if (glyphCount <= 0) {
            throw new NotSupportedException("OpenType font has an invalid glyph count.");
        }

        Dictionary<int, int> unicodeCMap = ReadUnicodeCMap(fontData, cmap);
        string resolvedFontName = SanitizePdfName(string.IsNullOrWhiteSpace(fontName) ? ReadPostScriptName(fontData, name) : fontName!);
        int cffLength = isOpenTypeCff ? tables["CFF "].Length : 0;
        bool hasGsubTable = tables.TryGetValue("GSUB", out TableRecord gsub);
        bool hasGposTable = tables.TryGetValue("GPOS", out TableRecord gpos);
        IReadOnlyList<string> gsubFeatures = hasGsubTable ? TryReadOpenTypeLayoutFeatureTags(fontData, gsub) : Array.Empty<string>();
        IReadOnlyList<string> gposFeatures = hasGposTable ? TryReadOpenTypeLayoutFeatureTags(fontData, gpos) : Array.Empty<string>();
        return new PdfOpenTypeFontInfo(
            resolvedFontName,
            FormatScalerType(scalerType),
            isOpenTypeCff,
            isTrueType,
            unitsPerEm,
            glyphCount,
            cffLength,
            unicodeCMap.Count,
            unicodeCMap,
            hasGsubTable,
            hasGposTable,
            gsubFeatures,
            gposFeatures);
    }

    /// <summary>
    /// Attempts to inspect OpenType metadata without throwing for unsupported or malformed font data.
    /// </summary>
    public static bool TryInspect(byte[] fontData, out PdfOpenTypeFontInfo? info, out string? error, string? fontName = null) {
        try {
            info = Inspect(fontData, fontName);
            error = null;
            return true;
        } catch (Exception exception) when (PdfFontDiagnostics.IsFontProgramException(exception)) {
            info = null;
            error = exception.Message;
            return false;
        }
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
                throw new NotSupportedException("OpenType font table offsets are too large.");
            }

            EnsureRange(data, (int)tableOffset, (int)tableLength);
            tables[tag] = new TableRecord((int)tableOffset, (int)tableLength);
        }

        return tables;
    }

    private static TableRecord GetTable(Dictionary<string, TableRecord> tables, string tag) {
        if (!tables.TryGetValue(tag, out TableRecord record)) {
            throw new NotSupportedException("OpenType font is missing required '" + tag + "' table.");
        }

        return record;
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
            throw new NotSupportedException("OpenType font does not contain a supported Unicode cmap subtable.");
        }

        int selectedFormat = ReadUInt16(data, offset);
        return selectedFormat == 12 ? ReadFormat12CMap(data, offset) : selectedFormat == 4 ? ReadFormat4CMap(data, offset) : ReadFormat0CMap(data, offset);
    }

    private static Dictionary<int, int> ReadFormat0CMap(byte[] data, int offset) {
        EnsureRange(data, offset, 262);
        var map = new Dictionary<int, int>();
        for (int code = 0; code < 256; code++) {
            int glyphId = data[offset + 6 + code];
            if (glyphId > 0) {
                map[code] = glyphId;
            }
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
        int processedMappings = 0;

        for (int segment = 0; segment < segCount; segment++) {
            int endCode = ReadUInt16(data, endCodeOffset + segment * 2);
            int startCode = ReadUInt16(data, startCodeOffset + segment * 2);
            int idDelta = ReadInt16(data, idDeltaOffset + segment * 2);
            int idRangeOffsetAddress = idRangeOffsetOffset + segment * 2;
            int idRangeOffset = ReadUInt16(data, idRangeOffsetAddress);
            if (startCode == 0xFFFF && endCode == 0xFFFF) {
                continue;
            }

            int mappingCount = endCode >= startCode ? endCode - startCode + 1 : 0;
            if (mappingCount == 0) {
                continue;
            }

            if (mappingCount > MaxUnicodeCMapMappings || processedMappings > MaxUnicodeCMapMappings - mappingCount) {
                throw new NotSupportedException("OpenType Unicode cmap mapping count exceeds supported limits.");
            }

            processedMappings += mappingCount;

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
            throw new NotSupportedException("OpenType format 12 cmap table is too large.");
        }

        EnsureRange(data, offset, (int)length);
        uint groupCount = ReadUInt32(data, offset + 12);
        if (groupCount > (uint)((length - 16) / 12)) {
            throw new NotSupportedException("OpenType format 12 cmap group count is invalid.");
        }

        var map = new Dictionary<int, int>();
        int processedMappings = 0;
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
            if (mappingCount > MaxUnicodeCMapMappings || processedMappings > MaxUnicodeCMapMappings - (int)mappingCount) {
                throw new NotSupportedException("OpenType Unicode cmap mapping count exceeds supported limits.");
            }

            processedMappings += (int)mappingCount;

            for (uint code = startCharCode; code <= endCharCode; code++) {
                uint glyph = startGlyphId + (code - startCharCode);
                if (glyph > 0 && glyph <= int.MaxValue) {
                    map[(int)code] = (int)glyph;
                }
            }
        }

        return map;
    }

    private static IReadOnlyList<string> TryReadOpenTypeLayoutFeatureTags(byte[] data, TableRecord layoutTable) {
        try {
            EnsureRange(data, layoutTable.Offset, Math.Min(layoutTable.Length, 10));
            int featureListOffset = ReadUInt16(data, layoutTable.Offset + 6);
            if (featureListOffset <= 0 || featureListOffset >= layoutTable.Length) {
                return Array.Empty<string>();
            }

            int featureListStart = checked(layoutTable.Offset + featureListOffset);
            EnsureRange(data, featureListStart, 2);
            int featureCount = ReadUInt16(data, featureListStart);
            int recordsStart = featureListStart + 2;
            int recordsLength = checked(featureCount * 6);
            if (featureCount < 0 ||
                recordsLength < 0 ||
                featureListOffset + 2 + recordsLength > layoutTable.Length) {
                return Array.Empty<string>();
            }

            EnsureRange(data, recordsStart, recordsLength);
            var tags = new List<string>();
            var seen = new HashSet<string>(StringComparer.Ordinal);
            for (int index = 0; index < featureCount; index++) {
                int record = recordsStart + index * 6;
                string tag = Encoding.ASCII.GetString(data, record, 4);
                if (IsReadableOpenTypeTag(tag) && seen.Add(tag)) {
                    tags.Add(tag);
                }
            }

            return tags;
        } catch (Exception exception) when (PdfFontDiagnostics.IsFontProgramException(exception)) {
            return Array.Empty<string>();
        }
    }

    private static bool IsReadableOpenTypeTag(string tag) {
        if (tag.Length != 4) {
            return false;
        }

        for (int index = 0; index < tag.Length; index++) {
            char ch = tag[index];
            if (ch < 0x20 || ch > 0x7E) {
                return false;
            }
        }

        return true;
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

    private static string FormatScalerType(uint scalerType) {
        if (scalerType == OpenTypeCffScalerType || scalerType == AppleTrueTypeScalerType) {
            return Encoding.ASCII.GetString(new[] {
                (byte)((scalerType >> 24) & 0xFF),
                (byte)((scalerType >> 16) & 0xFF),
                (byte)((scalerType >> 8) & 0xFF),
                (byte)(scalerType & 0xFF)
            });
        }

        return "0x" + scalerType.ToString("X8", CultureInfo.InvariantCulture);
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
            throw new NotSupportedException("OpenType font table data is truncated or invalid.");
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
