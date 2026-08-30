using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace OfficeIMO.Drawing;

/// <summary>Validated shared reader for sfnt table directories, cmap data, names, and horizontal metrics.</summary>
internal sealed class OfficeOpenTypeReader {
    private const int MaximumTables = 512;
    private readonly byte[] _data;
    private readonly Dictionary<uint, TableRecord> _tables;
    private readonly TableRecord _cmap;
    private readonly TableRecord _head;
    private readonly TableRecord _hhea;
    private readonly TableRecord _hmtx;
    private readonly TableRecord _maxp;
    private readonly TableRecord? _name;
    private readonly HashSet<int> _validFormat4Subtables;
    private readonly HashSet<int> _validFormat12Subtables;

    private OfficeOpenTypeReader(byte[] data, Dictionary<uint, TableRecord> tables) {
        _data = data;
        _tables = tables;
        _cmap = GetRequiredTable(Tag("cmap"));
        _head = GetRequiredTable(Tag("head"));
        _hhea = GetRequiredTable(Tag("hhea"));
        _hmtx = GetRequiredTable(Tag("hmtx"));
        _maxp = GetRequiredTable(Tag("maxp"));
        _name = tables.TryGetValue(Tag("name"), out TableRecord name) ? name : null;
        if (_head.Length < 54 || _hhea.Length < 36 || _maxp.Length < 6) {
            throw new InvalidDataException("The OpenType metrics tables are truncated.");
        }
        UnitsPerEm = ReadUInt16(_head.Offset + 18);
        Ascender = ReadInt16(_hhea.Offset + 4);
        Descender = ReadInt16(_hhea.Offset + 6);
        LineGap = ReadInt16(_hhea.Offset + 8);
        GlyphCount = ReadUInt16(_maxp.Offset + 4);
        HorizontalMetricCount = ReadUInt16(_hhea.Offset + 34);
        if (UnitsPerEm <= 0 || GlyphCount <= 0 || HorizontalMetricCount <= 0 || HorizontalMetricCount > GlyphCount) {
            throw new InvalidDataException("The OpenType font metrics are invalid.");
        }
        int requiredHmtxLength = checked(HorizontalMetricCount * 4 + (GlyphCount - HorizontalMetricCount) * 2);
        if (_hmtx.Length < requiredHmtxLength) throw new InvalidDataException("The OpenType hmtx table is truncated.");
        _validFormat4Subtables = OfficeOpenTypeCmap.CollectValidFormat4Subtables(
            _data,
            _cmap.Offset,
            _cmap.Length,
            OfficeOpenTypeCmap.MaximumSubtables);
        _validFormat12Subtables = OfficeOpenTypeCmap.CollectValidFormat12Subtables(
            _data,
            _cmap.Offset,
            _cmap.Length,
            OfficeOpenTypeCmap.MaximumSubtables,
            OfficeOpenTypeCmap.MaximumFormat12Groups);
    }

    internal byte[] Data => _data;
    internal int UnitsPerEm { get; }
    internal short Ascender { get; }
    internal short Descender { get; }
    internal short LineGap { get; }
    internal int GlyphCount { get; }
    internal int HorizontalMetricCount { get; }

    internal static OfficeOpenTypeReader? TryCreate(byte[] data) {
        if (data == null) throw new ArgumentNullException(nameof(data));
        try {
            if (data.Length < 12) return null;
            uint flavor = ReadUInt32(data, 0);
            if (flavor != 0x00010000 && flavor != 0x74727565 && flavor != 0x4F54544F) return null;
            int tableCount = ReadUInt16(data, 4);
            if (tableCount <= 0 || tableCount > MaximumTables || data.Length < 12 + tableCount * 16) return null;
            int directoryEnd = checked(12 + tableCount * 16);
            var tables = new Dictionary<uint, TableRecord>(tableCount);
            var ranges = new List<TableRecord>(tableCount);
            for (int index = 0; index < tableCount; index++) {
                int recordOffset = 12 + index * 16;
                uint tag = ReadUInt32(data, recordOffset);
                uint offsetValue = ReadUInt32(data, recordOffset + 8);
                uint lengthValue = ReadUInt32(data, recordOffset + 12);
                if (offsetValue > int.MaxValue || lengthValue > int.MaxValue) return null;
                int offset = (int)offsetValue;
                int length = (int)lengthValue;
                if (offset < 0 || length < 0 || offset > data.Length - length || tables.ContainsKey(tag) ||
                    (offset & 3) != 0 || length > 0 && offset < directoryEnd) return null;
                var record = new TableRecord(offset, length);
                if (length > 0) {
                    foreach (TableRecord range in ranges) {
                        if (offset < range.Offset + range.Length && range.Offset < offset + length) return null;
                    }
                    ranges.Add(record);
                }
                tables.Add(tag, record);
            }
            return new OfficeOpenTypeReader(data, tables);
        } catch (Exception exception) when (exception is InvalidDataException
                                            || exception is OverflowException
                                            || exception is ArgumentOutOfRangeException
                                            || exception is IndexOutOfRangeException) {
            return null;
        }
    }

    internal bool TryGetTable(string tag, out int offset, out int length) => TryGetTable(Tag(tag), out offset, out length);

    internal bool TryGetTable(uint tag, out int offset, out int length) {
        if (_tables.TryGetValue(tag, out TableRecord record)) {
            offset = record.Offset;
            length = record.Length;
            return true;
        }
        offset = 0;
        length = 0;
        return false;
    }

    internal int AdvanceWidth(int glyphId) {
        if (glyphId < 0 || glyphId >= GlyphCount) throw new ArgumentOutOfRangeException(nameof(glyphId));
        int metricIndex = Math.Min(glyphId, HorizontalMetricCount - 1);
        return ReadUInt16(_hmtx.Offset + metricIndex * 4);
    }

    internal int MapGlyph(int scalar) {
        if (scalar < 0 || scalar > 0x10FFFF) return 0;
        int table = SelectCmapSubtable(preferFormat12: scalar > 0xFFFF, out int cmapEnd);
        if (table < 0) return 0;
        int glyph = MapCmapSubtable(table, cmapEnd, scalar);
        return glyph != 0 ? glyph : MapFallbackCmapSubtable(table, cmapEnd, scalar);
    }

    internal bool HasGlyphs(string text) => OfficeOpenTypeCmap.HasGlyphs(
        text,
        MapGlyph,
        (scalar, selector) => OfficeOpenTypeCmap.SupportsVariationSequence(
            _data,
            _cmap.Offset,
            _cmap.Length,
            GlyphCount,
            scalar,
            selector,
            MapGlyph));

    private int MapCmapSubtable(int table, int cmapEnd, int scalar) {
        int format = ReadUInt16(table);
        return format == 12
            ? MapFormat12(table, cmapEnd, scalar)
            : format == 4 && scalar <= 0xFFFF
                ? MapFormat4(table, cmapEnd, scalar)
                : 0;
    }

    internal string? ReadDisplayName() => ReadFirstName(4) ?? ReadFirstName(1) ?? ReadFirstName(6) ?? ReadFirstName(2);

    internal ushort ReadUInt16(int offset) {
        EnsureAvailable(offset, 2);
        return ReadUInt16(_data, offset);
    }

    internal short ReadInt16(int offset) => unchecked((short)ReadUInt16(offset));

    internal uint ReadUInt32(int offset) {
        EnsureAvailable(offset, 4);
        return ReadUInt32(_data, offset);
    }

    internal int ReadInt32(int offset) => unchecked((int)ReadUInt32(offset));

    internal double ReadFixed16_16(int offset) => ReadInt32(offset) / 65536D;

    internal double ReadF2Dot14(int offset) => ReadInt16(offset) / 16384D;

    internal byte[] Slice(int offset, int length) {
        EnsureAvailable(offset, length);
        var result = new byte[length];
        Buffer.BlockCopy(_data, offset, result, 0, length);
        return result;
    }

    internal void EnsureAvailable(int offset, int length) {
        if (offset < 0 || length < 0 || offset > _data.Length - length) throw new InvalidDataException("The OpenType font data is truncated.");
    }

    private int SelectCmapSubtable(bool preferFormat12, out int cmapEnd) {
        cmapEnd = checked(_cmap.Offset + _cmap.Length);
        if (_cmap.Length < 4) return -1;
        int count = ReadUInt16(_cmap.Offset + 2);
        if (count <= 0 || count > OfficeOpenTypeCmap.MaximumSubtables || _cmap.Length < 4 + count * 8) return -1;
        int best = -1;
        int bestScore = int.MinValue;
        for (int index = 0; index < count; index++) {
            int record = _cmap.Offset + 4 + index * 8;
            int platform = ReadUInt16(record);
            int encoding = ReadUInt16(record + 2);
            uint relativeValue = ReadUInt32(record + 4);
            if (relativeValue > (uint)(_cmap.Length - 2)) continue;
            int subtable = _cmap.Offset + (int)relativeValue;
            if (subtable < _cmap.Offset || subtable > cmapEnd - 2) continue;
            int format = ReadUInt16(subtable);
            if (format != 4 && format != 12) continue;
            if (format == 4 && !_validFormat4Subtables.Contains(subtable)) continue;
            if (format == 12 && !_validFormat12Subtables.Contains(subtable)) continue;
            if (!OfficeOpenTypeCmap.IsUnicodeEncoding(platform, encoding)) continue;
            int score = OfficeOpenTypeCmap.ScoreSubtable(format, platform, encoding, preferFormat12);
            if (score > bestScore) {
                best = subtable;
                bestScore = score;
            }
        }
        return best;
    }

    private int MapFallbackCmapSubtable(int selectedTable, int cmapEnd, int scalar) {
        int count = ReadUInt16(_cmap.Offset + 2);
        int bestGlyph = 0;
        int bestScore = int.MinValue;
        for (int index = 0; index < count; index++) {
            int record = _cmap.Offset + 4 + index * 8;
            int platform = ReadUInt16(record);
            int encoding = ReadUInt16(record + 2);
            uint relativeValue = ReadUInt32(record + 4);
            if (relativeValue > (uint)(_cmap.Length - 2)) continue;
            int subtable = _cmap.Offset + (int)relativeValue;
            if (subtable == selectedTable || subtable < _cmap.Offset || subtable > cmapEnd - 2) continue;
            int format = ReadUInt16(subtable);
            if (format == 4 && (scalar > 0xFFFF || !_validFormat4Subtables.Contains(subtable))) continue;
            if (format == 12 && !_validFormat12Subtables.Contains(subtable)) continue;
            if (format != 4 && format != 12 || !OfficeOpenTypeCmap.IsUnicodeEncoding(platform, encoding)) continue;
            int score = OfficeOpenTypeCmap.ScoreSubtable(format, platform, encoding, preferFormat12: scalar > 0xFFFF);
            if (score <= bestScore) continue;
            int glyph = MapCmapSubtable(subtable, cmapEnd, scalar);
            if (glyph == 0) continue;
            bestGlyph = glyph;
            bestScore = score;
        }
        return bestGlyph;
    }

    private int MapFormat4(int table, int cmapEnd, int scalar) {
        if (!_validFormat4Subtables.Contains(table)) return 0;
        if (table < _cmap.Offset || table > cmapEnd - 14) return 0;
        int length = ReadUInt16(table + 2);
        int segmentCount = ReadUInt16(table + 6) / 2;
        if (length < 16 || segmentCount <= 0 || table > cmapEnd - length) return 0;
        int endCodes = table + 14;
        int startCodes = endCodes + segmentCount * 2 + 2;
        int deltas = startCodes + segmentCount * 2;
        int rangeOffsets = deltas + segmentCount * 2;
        if (rangeOffsets > table + length - segmentCount * 2) return 0;
        int low = 0;
        int high = segmentCount - 1;
        while (low <= high) {
            int index = low + (high - low) / 2;
            int end = ReadUInt16(endCodes + index * 2);
            int start = ReadUInt16(startCodes + index * 2);
            if (scalar < start) {
                high = index - 1;
                continue;
            }
            if (scalar > end) {
                low = index + 1;
                continue;
            }
            int delta = ReadInt16(deltas + index * 2);
            int rangeOffset = ReadUInt16(rangeOffsets + index * 2);
            if (rangeOffset == 0) return ValidateMappedGlyph(unchecked((ushort)(scalar + delta)));
            int glyphOffset = rangeOffsets + index * 2 + rangeOffset + (scalar - start) * 2;
            if (glyphOffset < table || glyphOffset > table + length - 2) return 0;
            int glyph = ReadUInt16(glyphOffset);
            return glyph == 0 ? 0 : ValidateMappedGlyph(unchecked((ushort)(glyph + delta)));
        }
        return 0;
    }

    private int MapFormat12(int table, int cmapEnd, int scalar) {
        if (!_validFormat12Subtables.Contains(table)) return 0;
        if (table < _cmap.Offset || table > cmapEnd - 16) return 0;
        uint lengthValue = ReadUInt32(table + 4);
        uint groupCount = ReadUInt32(table + 12);
        if (lengthValue > int.MaxValue || groupCount > OfficeOpenTypeCmap.MaximumFormat12Groups) return 0;
        int length = (int)lengthValue;
        if (length < 16 || table > cmapEnd - length || 16L + groupCount * 12L > length) return 0;
        int low = 0;
        int high = checked((int)groupCount - 1);
        while (low <= high) {
            int middle = low + ((high - low) / 2);
            int group = table + 16 + middle * 12;
            uint start = ReadUInt32(group);
            uint end = ReadUInt32(group + 4);
            if ((uint)scalar < start) high = middle - 1;
            else if ((uint)scalar > end) low = middle + 1;
            else {
                uint startGlyph = ReadUInt32(group + 8);
                ulong glyph = (ulong)startGlyph + (uint)scalar - start;
                return glyph < (ulong)GlyphCount ? checked((int)glyph) : 0;
            }
        }
        return 0;
    }

    private int ValidateMappedGlyph(int glyph) => glyph > 0 && glyph < GlyphCount ? glyph : 0;

    private string? ReadFirstName(ushort requestedNameId) {
        if (!_name.HasValue) return null;
        TableRecord name = _name.Value;
        if (name.Length < 6) return null;
        int count = ReadUInt16(name.Offset + 2);
        int storage = name.Offset + ReadUInt16(name.Offset + 4);
        if (count < 0 || name.Length < 6 + count * 12 || storage < name.Offset || storage > name.Offset + name.Length) return null;
        string? fallback = null;
        for (int index = 0; index < count; index++) {
            int record = name.Offset + 6 + index * 12;
            int platform = ReadUInt16(record);
            int language = ReadUInt16(record + 4);
            int nameId = ReadUInt16(record + 6);
            int length = ReadUInt16(record + 8);
            int offset = ReadUInt16(record + 10);
            if (nameId != requestedNameId || storage + offset < storage || storage + offset > name.Offset + name.Length - length) continue;
            string value;
            if (platform == 0 || platform == 3) {
                if ((length & 1) != 0) continue;
                var chars = new char[length / 2];
                for (int character = 0; character < chars.Length; character++) chars[character] = (char)ReadUInt16(storage + offset + character * 2);
                value = new string(chars);
            } else {
                value = Encoding.ASCII.GetString(_data, storage + offset, length);
            }
            value = value.Trim('\0', ' ');
            if (value.Length == 0) continue;
            if (language == 0x0409 || language == 0) return value;
            fallback ??= value;
        }
        return fallback;
    }

    private TableRecord GetRequiredTable(uint tag) {
        if (!_tables.TryGetValue(tag, out TableRecord record)) throw new InvalidDataException("The OpenType font is missing a required table.");
        return record;
    }

    private static ushort ReadUInt16(byte[] data, int offset) => unchecked((ushort)((data[offset] << 8) | data[offset + 1]));

    private static uint ReadUInt32(byte[] data, int offset) => unchecked(((uint)data[offset] << 24)
        | ((uint)data[offset + 1] << 16)
        | ((uint)data[offset + 2] << 8)
        | data[offset + 3]);

    internal static uint Tag(string value) {
        if (value == null || value.Length != 4) throw new ArgumentException("OpenType tags must contain four characters.", nameof(value));
        return ((uint)value[0] << 24) | ((uint)value[1] << 16) | ((uint)value[2] << 8) | value[3];
    }

    private readonly struct TableRecord {
        internal TableRecord(int offset, int length) {
            Offset = offset;
            Length = length;
        }

        internal int Offset { get; }
        internal int Length { get; }
    }
}
