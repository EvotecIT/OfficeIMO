using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeTrueTypeFont {
    private readonly FixedGlyphTables _fixedGlyphTables;

    private sealed class FixedGlyphTables {
        internal readonly int VerticalHeader, VerticalHeaderLength, VerticalMetrics, VerticalMetricsLength;
        internal readonly int Os2, Os2Length, GlyphLength, LocationLength, HorizontalLength;
        internal FixedGlyphTables(Dictionary<string, int> offsets, IReadOnlyDictionary<string, int> lengths) {
            VerticalHeader = offsets.TryGetValue("vhea", out int vh) ? vh : -1;
            VerticalHeaderLength = lengths.TryGetValue("vhea", out int vl) ? vl : 0;
            VerticalMetrics = offsets.TryGetValue("vmtx", out int vm) ? vm : -1;
            VerticalMetricsLength = lengths.TryGetValue("vmtx", out int ml) ? ml : 0;
            Os2 = offsets.TryGetValue("OS/2", out int os) ? os : -1;
            Os2Length = lengths.TryGetValue("OS/2", out int ol) ? ol : 0;
            GlyphLength = lengths["glyf"];
            LocationLength = lengths["loca"];
            HorizontalLength = lengths["hmtx"];
        }
    }

    /// <summary>Reads the top-center origin and vertical advance for fixed, default-instance TrueType glyph placement.</summary>
    /// <remarks>Uses vmtx when present, otherwise OS/2 typographic metrics or hhea, as required by XPS sideways layout.</remarks>
    internal double FixedGlyphVerticalMetrics(int glyphId, double fontSize, out double originX, out double originY) {
        if (glyphId < 0 || glyphId >= _numGlyphs) throw new ArgumentOutOfRangeException(nameof(glyphId));
        var tables = _fixedGlyphTables;
        if (_numHMetrics == 0 || _numHMetrics > _numGlyphs || tables.HorizontalLength < _numHMetrics * 4 + (_numGlyphs - _numHMetrics) * 2)
            throw new InvalidDataException("Invalid horizontal glyph metrics.");
        double scale = ScaleFor(fontSize);
        originX = FixedGlyphAdvance(glyphId, fontSize) / 2;
        int top, advance;
        if (tables.VerticalMetrics >= 0) {
            if (tables.VerticalHeaderLength < 36) throw new InvalidDataException("Invalid vertical metrics header.");
            int count = ReadUInt16(_data, tables.VerticalHeader + 34);
            if (count == 0 || count > _numGlyphs || tables.VerticalMetricsLength < count * 4 + (_numGlyphs - count) * 2)
                throw new InvalidDataException("Invalid vertical glyph metrics.");
            int metric = tables.VerticalMetrics + Math.Min(glyphId, count - 1) * 4;
            advance = ReadUInt16(_data, metric);
            int bearing = glyphId < count ? metric + 2 : tables.VerticalMetrics + count * 4 + (glyphId - count) * 2;
            int stride = _indexToLocFormat == 0 ? 2 : 4;
            if ((_indexToLocFormat != 0 && _indexToLocFormat != 1) || tables.LocationLength < (glyphId + 2) * stride)
                throw new InvalidDataException("Invalid glyph location table.");
            int start = GlyphOffset((ushort)glyphId), end = GlyphOffset((ushort)(glyphId + 1));
            if (start < 0 || end < start || end > tables.GlyphLength || (end != start && end - start < 10))
                throw new InvalidDataException("Invalid vertical glyph bounds.");
            int ymax = start == end ? 0 : ReadInt16(_data, _glyf + start + 8);
            top = ymax + ReadInt16(_data, bearing);
        } else if (tables.Os2Length >= 78) {
            top = ReadInt16(_data, tables.Os2 + 68);
            advance = top + Math.Abs((int)ReadInt16(_data, tables.Os2 + 70));
        } else {
            top = _ascender;
            advance = top + Math.Abs(_descender);
        }
        if (advance < 0) throw new InvalidDataException("Invalid vertical glyph advance.");
        originY = top * scale;
        return advance * scale;
    }
    /// <summary>Reads the native horizontal advance for an already resolved glyph index.</summary>
    internal double FixedGlyphAdvance(int glyphId, double fontSize) {
        if (glyphId < 0 || glyphId >= _numGlyphs) throw new ArgumentOutOfRangeException(nameof(glyphId));
        return AdvanceWidth((ushort)glyphId) * ScaleFor(fontSize);
    }
    /// <summary>Projects an explicitly positioned glyph without Unicode remapping or shaping.</summary>
    internal List<List<OfficePoint>> FixedGlyphContours(int glyphId, double fontSize, double x, double baseline,
        int maximumPoints, CancellationToken cancellationToken) {
        if (glyphId < 0 || glyphId >= _numGlyphs) throw new ArgumentOutOfRangeException(nameof(glyphId));
        if (maximumPoints <= 0) throw new ArgumentOutOfRangeException(nameof(maximumPoints));
        double scale = ScaleFor(fontSize);
        int count = 0;
        return ReadGlyphContours((ushort)glyphId, new FontTransform(scale, 0, 0, -scale, x, baseline), 0,
            _variations?.CreateWorkBudget(), maximumPoints, ref count, cancellationToken, attachmentPoints: null);
    }
}
