namespace OfficeIMO.Drawing;

/// <summary>A bitmap color glyph and its placement in font units, y up, relative to the glyph origin.</summary>
internal sealed class OfficeBitmapGlyphImage {
    internal OfficeBitmapGlyphImage(byte[] png, double left, double bottom, double width, double height) {
        Png = png; Left = left; Bottom = bottom; Width = width; Height = height;
    }

    /// <summary>The PNG image data.</summary>
    internal byte[] Png { get; }
    internal double Left { get; }
    internal double Bottom { get; }
    internal double Width { get; }
    internal double Height { get; }
}

/// <summary>
/// Bounded reader for PNG color glyphs in CBDT/CBLC and sbix tables (OpenType specification),
/// as Noto Color Emoji and Apple Color Emoji use. The largest strike is selected.
/// </summary>
internal sealed class OfficeOpenTypeBitmapGlyphs {
    private const int MaximumStrikes = 256;
    private const int MaximumIndexSubtables = 65536;
    private const int MaximumImageBytes = 16 * 1024 * 1024;
    private const int MaximumCachedGlyphs = 4096;
    private static readonly byte[] PngSignature = { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };

    private readonly OfficeOpenTypeReader _reader;
    private readonly int _unitsPerEm;
    private readonly int _glyphCount;
    private readonly int _cblc, _cblcEnd, _cbdt, _cbdtEnd, _cbdtStrike = -1;
    private readonly int _sbix, _sbixEnd, _sbixStrike = -1;
    private readonly Dictionary<int, OfficeBitmapGlyphImage?> _cache = new();

    private OfficeOpenTypeBitmapGlyphs(OfficeOpenTypeReader reader, int unitsPerEm) {
        _reader = reader;
        _unitsPerEm = unitsPerEm;
        _glyphCount = reader.GlyphCount;
        if (reader.TryGetTable("CBLC", out _cblc, out int cblcLength) && reader.TryGetTable("CBDT", out _cbdt, out int cbdtLength)
            && cblcLength >= 8 && cbdtLength >= 4) {
            _cblcEnd = checked(_cblc + cblcLength);
            _cbdtEnd = checked(_cbdt + cbdtLength);
            _cbdtStrike = SelectBitmapSize();
        }
        if (reader.TryGetTable("sbix", out _sbix, out int sbixLength) && sbixLength >= 8) {
            _sbixEnd = checked(_sbix + sbixLength);
            _sbixStrike = SelectSbixStrike();
        }
    }

    internal static OfficeOpenTypeBitmapGlyphs? TryCreate(OfficeOpenTypeReader reader, int unitsPerEm) {
        ArgumentNullException.ThrowIfNull(reader);
        if (unitsPerEm <= 0) return null;
        try {
            var glyphs = new OfficeOpenTypeBitmapGlyphs(reader, unitsPerEm);
            return glyphs._cbdtStrike >= 0 || glyphs._sbixStrike >= 0 ? glyphs : null;
        } catch (Exception exception) when (IsMalformed(exception)) {
            return null;
        }
    }

    /// <summary>Returns the PNG image of a glyph from the selected strike, or false when it has none.</summary>
    internal bool TryGetImage(int glyphId, out OfficeBitmapGlyphImage? image) {
        lock (_cache) {
            if (_cache.TryGetValue(glyphId, out image)) return image != null;
        }
        try {
            image = glyphId <= 0 || glyphId >= _glyphCount ? null : ReadCbdt(glyphId) ?? ReadSbix(glyphId, allowDuplicate: true);
        } catch (Exception exception) when (IsMalformed(exception)) {
            image = null;
        }
        lock (_cache) {
            if (_cache.Count >= MaximumCachedGlyphs) _cache.Clear();
            _cache[glyphId] = image;
        }
        return image != null;
    }

    private int SelectBitmapSize() {
        int major = U16(_cblc, _cblcEnd);
        if (major != 2 && major != 3) return -1;
        uint count = U32(_cblc + 4, _cblcEnd);
        if (count == 0 || count > MaximumStrikes) return -1;
        int best = -1, bestPpem = -1;
        for (int index = 0; index < (int)count; index++) {
            int record = checked(_cblc + 8 + index * 48);
            Ensure(record, 48, _cblcEnd);
            int ppem = _reader.Data[record + 45];
            if (ppem > bestPpem) { bestPpem = ppem; best = record; }
        }
        return best;
    }

    private OfficeBitmapGlyphImage? ReadCbdt(int glyphId) {
        if (_cbdtStrike < 0) return null;
        int record = _cbdtStrike;
        int startGlyph = U16(record + 40, _cblcEnd), endGlyph = U16(record + 42, _cblcEnd);
        if (glyphId < startGlyph || glyphId > endGlyph) return null;
        int ppem = _reader.Data[record + 45];
        if (ppem == 0) return null;
        uint arrayOffset = U32(record, _cblcEnd);
        uint subtableCount = U32(record + 8, _cblcEnd);
        if (subtableCount > MaximumIndexSubtables || arrayOffset > int.MaxValue) return null;
        int array = checked(_cblc + (int)arrayOffset);
        for (int index = 0; index < (int)subtableCount; index++) {
            int entry = checked(array + index * 8);
            int first = U16(entry, _cblcEnd), last = U16(entry + 2, _cblcEnd);
            if (glyphId < first || glyphId > last) continue;
            uint additional = U32(entry + 4, _cblcEnd);
            if (additional > int.MaxValue) return null;
            int subtable = checked(array + (int)additional);
            int indexFormat = U16(subtable, _cblcEnd), imageFormat = U16(subtable + 2, _cblcEnd);
            uint dataOffset = U32(subtable + 4, _cblcEnd);
            if (dataOffset > int.MaxValue) return null;
            int imageData = checked(_cbdt + (int)dataOffset);
            int start, length;
            int metrics = -1;
            switch (indexFormat) {
                case 1: {
                    int offsets = subtable + 8, position = glyphId - first;
                    uint a = U32(offsets + position * 4, _cblcEnd), b = U32(offsets + position * 4 + 4, _cblcEnd);
                    if (b < a || b > int.MaxValue) return null;
                    start = checked(imageData + (int)a); length = (int)(b - a);
                    break;
                }
                case 2: {
                    uint size = U32(subtable + 8, _cblcEnd);
                    if (size > MaximumImageBytes) return null;
                    metrics = subtable + 12;
                    start = checked(imageData + (glyphId - first) * (int)size); length = (int)size;
                    break;
                }
                case 3: {
                    int offsets = subtable + 8, position = glyphId - first;
                    int a = U16(offsets + position * 2, _cblcEnd), b = U16(offsets + position * 2 + 2, _cblcEnd);
                    if (b < a) return null;
                    start = checked(imageData + a); length = b - a;
                    break;
                }
                case 4: {
                    uint glyphs = U32(subtable + 8, _cblcEnd);
                    if (glyphs > 65536) return null;
                    start = -1; length = 0;
                    for (int item = 0; item < (int)glyphs; item++) {
                        int pair = subtable + 12 + item * 4;
                        if (U16(pair, _cblcEnd) != glyphId) continue;
                        int a = U16(pair + 2, _cblcEnd), b = U16(pair + 6, _cblcEnd);
                        if (b < a) return null;
                        start = checked(imageData + a); length = b - a;
                        break;
                    }
                    if (start < 0) return null;
                    break;
                }
                case 5: {
                    uint size = U32(subtable + 8, _cblcEnd), glyphs = U32(subtable + 20, _cblcEnd);
                    if (size > MaximumImageBytes || glyphs > 65536) return null;
                    metrics = subtable + 12;
                    start = -1; length = (int)size;
                    for (int item = 0; item < (int)glyphs; item++) {
                        if (U16(subtable + 24 + item * 2, _cblcEnd) != glyphId) continue;
                        start = checked(imageData + item * (int)size);
                        break;
                    }
                    if (start < 0) return null;
                    break;
                }
                default:
                    return null;
            }
            Ensure(start, length, _cbdtEnd);
            return ReadCbdtImage(imageFormat, start, length, metrics, ppem);
        }
        return null;
    }

    // Image formats 17 and 18 carry their own small or big metrics; format 19 uses the index's big metrics.
    private OfficeBitmapGlyphImage? ReadCbdtImage(int format, int start, int length, int indexMetrics, int ppem) {
        int width, height, bearingX, bearingY, data;
        switch (format) {
            case 17:
                Ensure(start, 9, start + length);
                height = _reader.Data[start]; width = _reader.Data[start + 1];
                bearingX = (sbyte)_reader.Data[start + 2]; bearingY = (sbyte)_reader.Data[start + 3];
                data = start + 5;
                break;
            case 18:
                Ensure(start, 12, start + length);
                height = _reader.Data[start]; width = _reader.Data[start + 1];
                bearingX = (sbyte)_reader.Data[start + 2]; bearingY = (sbyte)_reader.Data[start + 3];
                data = start + 8;
                break;
            case 19:
                if (indexMetrics < 0) return null;
                Ensure(indexMetrics, 8, _cblcEnd);
                height = _reader.Data[indexMetrics]; width = _reader.Data[indexMetrics + 1];
                bearingX = (sbyte)_reader.Data[indexMetrics + 2]; bearingY = (sbyte)_reader.Data[indexMetrics + 3];
                data = start;
                break;
            default:
                return null;
        }
        uint pngLength = U32(data, start + length);
        if (pngLength == 0 || pngLength > MaximumImageBytes || pngLength > (uint)(start + length - data - 4)) return null;
        byte[]? png = CopyPng(data + 4, (int)pngLength);
        if (png == null || width == 0 || height == 0) return null;
        double scale = (double)_unitsPerEm / ppem;
        return new OfficeBitmapGlyphImage(png, bearingX * scale, (bearingY - height) * scale, width * scale, height * scale);
    }

    private int SelectSbixStrike() {
        uint count = U32(_sbix + 4, _sbixEnd);
        if (count == 0 || count > MaximumStrikes) return -1;
        int best = -1, bestPpem = -1;
        for (int index = 0; index < (int)count; index++) {
            uint offset = U32(_sbix + 8 + index * 4, _sbixEnd);
            if (offset > int.MaxValue) continue;
            int strike = checked(_sbix + (int)offset);
            int ppem = U16(strike, _sbixEnd);
            if (ppem > bestPpem) { bestPpem = ppem; best = strike; }
        }
        return bestPpem > 0 ? best : -1;
    }

    private OfficeBitmapGlyphImage? ReadSbix(int glyphId, bool allowDuplicate) {
        if (_sbixStrike < 0) return null;
        int strike = _sbixStrike;
        int ppem = U16(strike, _sbixEnd);
        uint a = U32(strike + 4 + glyphId * 4, _sbixEnd), b = U32(strike + 8 + glyphId * 4, _sbixEnd);
        if (b <= a || b - a < 8 || b > int.MaxValue) return null;
        int start = checked(strike + (int)a), end = checked(strike + (int)b);
        Ensure(start, end - start, _sbixEnd);
        short originX = unchecked((short)U16(start, end)), originY = unchecked((short)U16(start + 2, end));
        uint type = U32(start + 4, end);
        if (type == 0x64757065 && allowDuplicate) { // 'dupe'
            int target = U16(start + 8, end);
            return target != glyphId && target > 0 && target < _glyphCount ? ReadSbix(target, allowDuplicate: false) : null;
        }
        if (type != 0x706E6720) return null; // 'png '
        byte[]? png = CopyPng(start + 8, end - start - 8);
        if (png == null || png.Length < 24) return null;
        int width = (int)ReadBigEndian32(png, 16), height = (int)ReadBigEndian32(png, 20);
        if (width <= 0 || height <= 0) return null;
        double scale = (double)_unitsPerEm / ppem;
        return new OfficeBitmapGlyphImage(png, originX * scale, originY * scale, width * scale, height * scale);
    }

    private byte[]? CopyPng(int start, int length) {
        if (length < PngSignature.Length || length > MaximumImageBytes) return null;
        Ensure(start, length, _reader.Data.Length);
        for (int index = 0; index < PngSignature.Length; index++) if (_reader.Data[start + index] != PngSignature[index]) return null;
        var png = new byte[length];
        Array.Copy(_reader.Data, start, png, 0, length);
        return png;
    }

    private int U16(int offset, int end) {
        Ensure(offset, 2, end);
        return _reader.ReadUInt16(offset);
    }

    private uint U32(int offset, int end) {
        Ensure(offset, 4, end);
        return _reader.ReadUInt32(offset);
    }

    private static uint ReadBigEndian32(byte[] data, int offset) =>
        (uint)(data[offset] << 24 | data[offset + 1] << 16 | data[offset + 2] << 8 | data[offset + 3]);

    private static void Ensure(int offset, int length, int end) {
        if (offset < 0 || length < 0 || offset > end - length) throw new InvalidDataException("A bitmap glyph table record is outside its table.");
    }

    private static bool IsMalformed(Exception exception) => exception is InvalidDataException or OverflowException
        or ArgumentOutOfRangeException or IndexOutOfRangeException;
}
