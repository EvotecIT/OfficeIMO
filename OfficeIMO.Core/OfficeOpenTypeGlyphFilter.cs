using System;
using System.Globalization;
using System.IO;

namespace OfficeIMO.Drawing;

/// <summary>Reads GDEF glyph classes and the filtering rules shared by GSUB and GPOS lookups.</summary>
internal sealed class OfficeOpenTypeGlyphFilter {
    private readonly byte[] _data;
    private readonly int _start, _end, _classes, _attachmentClasses, _markSets;

    private OfficeOpenTypeGlyphFilter(byte[] data, OfficeOpenTypeReader reader) {
        _data = data;
        if (!reader.TryGetTable("GDEF", out _start, out int length)) { _start = _end = -1; return; }
        _end = checked(_start + length);
        Ensure(_start, 12);
        if (U16(_start) != 1) throw new InvalidDataException("Unsupported OpenType GDEF version.");
        int minor = U16(_start + 2);
        _classes = Optional(_start, U16(_start + 4), 4);
        _attachmentClasses = Optional(_start, U16(_start + 10), 4);
        if (minor >= 2) { Ensure(_start, 14); _markSets = Optional(_start, U16(_start + 12), 4); }
    }

    internal static OfficeOpenTypeGlyphFilter? TryCreate(byte[] data) {
        if (data == null) throw new ArgumentNullException(nameof(data));
        try {
            OfficeOpenTypeReader? reader = OfficeOpenTypeReader.TryCreate(data);
            return reader == null ? null : new OfficeOpenTypeGlyphFilter(data, reader);
        } catch (Exception exception) when (exception is InvalidDataException or OverflowException
            or ArgumentOutOfRangeException or IndexOutOfRangeException) { return null; }
    }

    /// <summary>Uses Unicode categories only when the font supplies no glyph-class definition.</summary>
    internal int GetGlyphClass(int glyphId, int scalar, bool isLigature, Action charge) {
        if ((uint)glyphId > ushort.MaxValue) throw new InvalidDataException("OpenType glyph is out of range.");
        if (_classes != 0) return Class(_classes, glyphId, charge);
        if (IsUnicodeMark(scalar)) return 3;
        return isLigature ? 2 : 1;
    }

    internal bool SupportsFlags(ushort flags, int markFilteringSet) {
        if ((flags & 0x00E0) != 0) return false;
        if ((flags & 0xFF00) != 0 && _attachmentClasses == 0) return false;
        if ((flags & 0x0010) == 0) return true;
        if (_markSets == 0 || U16(_markSets) != 1) return false;
        int count = U16(_markSets + 2);
        Ensure(_markSets + 4, checked(count * 4));
        return markFilteringSet >= 0 && markFilteringSet < count;
    }

    internal bool IsIgnored(OfficeOpenTypeSubstitution.GlyphToken token, ushort flags,
        int markFilteringSet, Action charge) {
        charge();
        if (!SupportsFlags(flags, markFilteringSet)) throw new InvalidDataException("Unsupported OpenType glyph-filter flags.");
        int glyphClass = GetGlyphClass(token.GlyphId, token.Scalar, token.LigatureComponentTextIndexes?.Count > 1, charge);
        if (glyphClass == 1 && (flags & 0x0002) != 0 || glyphClass == 2 && (flags & 0x0004) != 0) return true;
        if (glyphClass != 3) return false;
        if ((flags & 0x0008) != 0) return true;
        // UseMarkFilteringSet takes precedence over MarkAttachmentType.
        if ((flags & 0x0010) != 0) {
            int coverage = Relative(_markSets, U32(_markSets + 4 + markFilteringSet * 4), 4);
            return Coverage(coverage, token.GlyphId, charge) < 0;
        }
        int attachmentType = flags >> 8;
        return attachmentType != 0 && Class(_attachmentClasses, token.GlyphId, charge) != attachmentType;
    }

    internal static bool IsUnicodeMark(int scalar) {
        if (scalar < 0 || scalar > 0x10FFFF || scalar is >= 0xD800 and <= 0xDFFF) return false;
        UnicodeCategory category = scalar <= char.MaxValue ? char.GetUnicodeCategory((char)scalar)
            : CharUnicodeInfo.GetUnicodeCategory(char.ConvertFromUtf32(scalar), 0);
        return category is UnicodeCategory.NonSpacingMark or UnicodeCategory.SpacingCombiningMark or UnicodeCategory.EnclosingMark;
    }

    private int Class(int table, int glyph, Action charge) {
        Ensure(table, 4);
        int format = U16(table);
        if (format == 1) {
            Ensure(table, 6);
            int index = glyph - U16(table + 2), count = U16(table + 4);
            Ensure(table + 6, checked(count * 2));
            charge();
            return index < 0 || index >= count ? 0 : U16(table + 6 + index * 2);
        }
        if (format != 2) throw new InvalidDataException("Unsupported OpenType class-definition format.");
        int ranges = U16(table + 2);
        Ensure(table + 4, checked(ranges * 6));
        int low = 0, high = ranges - 1;
        while (low <= high) {
            charge();
            int middle = low + ((high - low) / 2), record = table + 4 + middle * 6;
            if (glyph < U16(record)) high = middle - 1;
            else if (glyph > U16(record + 2)) low = middle + 1;
            else return U16(record + 4);
        }
        return 0;
    }

    private int Coverage(int table, int glyph, Action charge) {
        Ensure(table, 4);
        int format = U16(table), count = U16(table + 2);
        int stride = format == 1 ? 2 : format == 2 ? 6 : 0;
        if (stride == 0) throw new InvalidDataException("Unsupported OpenType coverage format.");
        Ensure(table + 4, checked(count * stride));
        int low = 0, high = count - 1;
        while (low <= high) {
            charge();
            int middle = low + ((high - low) / 2), record = table + 4 + middle * stride;
            int first = U16(record), last = format == 1 ? first : U16(record + 2);
            if (glyph < first) high = middle - 1;
            else if (glyph > last) low = middle + 1;
            else return format == 1 ? middle : checked(U16(record + 4) + glyph - first);
        }
        return -1;
    }

    private int Optional(int table, int relative, int size) => relative == 0 ? 0 : Relative(table, (uint)relative, size);
    private int Relative(int table, uint relative, int size) {
        if (relative > int.MaxValue) throw new InvalidDataException("OpenType relative offset is out of range.");
        int result = checked(table + (int)relative); Ensure(result, size); return result;
    }
    private void Ensure(int offset, int size) {
        if (size < 0 || offset < _start || offset > _end - size)
            throw new InvalidDataException("OpenType GDEF data is truncated.");
    }
    private ushort U16(int offset) { Ensure(offset, 2); return (ushort)((_data[offset] << 8) | _data[offset + 1]); }
    private uint U32(int offset) { Ensure(offset, 4); return ((uint)_data[offset] << 24) | ((uint)_data[offset + 1] << 16) | ((uint)_data[offset + 2] << 8) | _data[offset + 3]; }
}
