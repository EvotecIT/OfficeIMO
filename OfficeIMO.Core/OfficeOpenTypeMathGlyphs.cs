using System;
using System.Collections.Generic;
using System.IO;

namespace OfficeIMO.Drawing;

/// <summary>Table-local bounded reader for static MATH constructions and glyph placement.</summary>
internal sealed class OfficeOpenTypeMathGlyphs {
    private readonly byte[] _data;
    private readonly int _start, _length, _glyphCount;
    private int _records;
    private const int MaximumRecords = 65536;
    private const int MaximumConstructionRecords = 256;

    private OfficeOpenTypeMathGlyphs(byte[] data, int start, int length, int glyphCount) {
        _data = data; _start = start; _length = length; _glyphCount = glyphCount;
    }

    internal static OfficeMathGlyphData? TryRead(byte[] data, int offset, int length, int glyphCount) {
        if (offset < 0 || length < 10 || offset > data.Length - length || glyphCount <= 0) return null;
        try { return new OfficeOpenTypeMathGlyphs(data, offset, length, glyphCount).Read(); }
        catch (InvalidDataException) { return null; }
    }

    private OfficeMathGlyphData Read() {
        if (U16(0) != 1 || U16(2) != 0) throw Invalid();
        var result = new OfficeMathGlyphData();
        int variants = Offset(0, 8, 10, optional: true);
        if (variants != 0) {
            result.MinimumOverlap = U16(variants);
            int vertical = U16(variants + 6), horizontal = U16(variants + 8);
            Require(variants, 10 + (vertical + horizontal) * 2);
            ReadConstructions(result.Vertical, variants, vertical, variants + 10, variants + 2);
            ReadConstructions(result.Horizontal, variants, horizontal, variants + 10 + vertical * 2, variants + 4);
        }
        int info = Offset(0, 6, 8, optional: true);
        if (info != 0) {
            ReadValues(result.Italics, Offset(info, info, 4, optional: true));
            ReadValues(result.Accents, Offset(info, info + 2, 4, optional: true));
            int kern = Offset(info, info + 6, 4, optional: true);
            if (kern != 0) {
                int count = U16(kern + 2);
                Require(kern, 4 + count * 8);
                int[] coverage = Coverage(Offset(kern, kern, 4), count);
                for (int i = 0; i < count; i++) {
                    var corners = new OfficeMathKernTable?[4];
                    for (int corner = 0; corner < 4; corner++) {
                        int table = Offset(kern, kern + 4 + i * 8 + corner * 2, 2, optional: true);
                        if (table == 0) continue;
                        int heights = U16(table); Charge(heights + 1);
                        Require(table, 2 + (heights * 2 + 1) * 4);
                        var h = new int[heights]; var v = new int[heights + 1];
                        for (int j = 0; j < heights; j++) {
                            h[j] = I16(table + 2 + j * 4);
                            if (j > 0 && h[j] <= h[j - 1]) throw Invalid();
                        }
                        for (int j = 0; j <= heights; j++) v[j] = I16(table + 2 + heights * 4 + j * 4);
                        corners[corner] = new OfficeMathKernTable(h, v);
                    }
                    result.Kerns.Add(coverage[i], corners);
                }
            }
        }
        return result;
    }

    private void ReadValues(Dictionary<int, int> target, int table) {
        if (table == 0) return;
        int count = U16(table + 2); Charge(count);
        Require(table, 4 + count * 4);
        int[] coverage = Coverage(Offset(table, table, 4), count);
        for (int i = 0; i < count; i++) target.Add(coverage[i], I16(table + 4 + i * 4));
    }

    private void ReadConstructions(Dictionary<int, OfficeMathGlyphConstruction> target, int table,
        int count, int array, int coverageField) {
        if (count == 0) return;
        int[] coverage = Coverage(Offset(table, coverageField, 4), count);
        for (int i = 0; i < count; i++) {
            int construction = Offset(table, array + i * 2, 4);
            int variants = U16(construction + 2);
            if (variants > MaximumConstructionRecords) throw Invalid();
            Charge(variants); Require(construction, 4 + variants * 4);
            var v = new OfficeMathGlyphVariant[variants];
            for (int j = 0; j < variants; j++) {
                int glyph = Glyph(construction + 4 + j * 4), advance = U16(construction + 6 + j * 4);
                if (advance == 0 || j > 0 && advance < v[j - 1].Advance) throw Invalid();
                v[j] = new OfficeMathGlyphVariant(glyph, advance);
            }
            int assembly = Offset(construction, construction, 6, optional: true);
            var parts = Array.Empty<OfficeMathGlyphPart>(); int italic = 0;
            if (assembly != 0) {
                italic = I16(assembly);
                int partCount = U16(assembly + 4);
                if (partCount == 0 || partCount > MaximumConstructionRecords) throw Invalid();
                Charge(partCount); Require(assembly, 6 + partCount * 10);
                parts = new OfficeMathGlyphPart[partCount];
                for (int j = 0; j < partCount; j++) {
                    int p = assembly + 6 + j * 10;
                    int glyph = Glyph(p), start = U16(p + 2), end = U16(p + 4), advance = U16(p + 6), flags = U16(p + 8);
                    if (advance == 0 || (flags & ~1) != 0) throw Invalid();
                    parts[j] = new OfficeMathGlyphPart(glyph, start, end, advance, (flags & 1) != 0);
                }
            }
            target.Add(coverage[i], new OfficeMathGlyphConstruction(v, parts, italic));
        }
    }

    private int[] Coverage(int table, int expected) {
        Charge(expected);
        var glyphs = new int[expected];
        int format = U16(table), count = U16(table + 2);
        if (format == 1) {
            if (count != expected) throw Invalid(); Require(table, 4 + count * 2);
            for (int i = 0; i < count; i++) glyphs[i] = Glyph(table + 4 + i * 2);
        } else if (format == 2) {
            Require(table, 4 + count * 6); int index = 0;
            for (int i = 0; i < count; i++) {
                int range = table + 4 + i * 6, first = Glyph(range), last = Glyph(range + 2);
                if (last < first || U16(range + 4) != index || last - first + 1 > expected - index) throw Invalid();
                for (int g = first; g <= last; g++) glyphs[index++] = g;
            }
            if (index != expected) throw Invalid();
        } else throw Invalid();
        for (int i = 1; i < glyphs.Length; i++) if (glyphs[i] <= glyphs[i - 1]) throw Invalid();
        return glyphs;
    }

    private int Glyph(int offset) { int glyph = U16(offset); if (glyph == 0 || glyph >= _glyphCount) throw Invalid(); return glyph; }
    private int Offset(int parent, int field, int minimum, bool optional = false) {
        int relative = U16(field); if (relative == 0 && optional) return 0;
        if (relative == 0) throw Invalid(); int table = parent + relative; Require(table, minimum); return table;
    }
    private int U16(int offset) { Require(offset, 2); return (_data[_start + offset] << 8) | _data[_start + offset + 1]; }
    private int I16(int offset) => unchecked((short)U16(offset));
    private void Require(int offset, int count) { if (offset < 0 || count < 0 || offset > _length - count) throw Invalid(); }
    private void Charge(int count) { if (count > MaximumRecords - _records) throw Invalid(); _records += count; }
    private static InvalidDataException Invalid() => new("Invalid or unsupported bounded MATH glyph data.");
}
