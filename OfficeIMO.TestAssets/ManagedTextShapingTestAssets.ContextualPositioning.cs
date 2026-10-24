using System;
using System.Collections.Generic;

namespace OfficeIMO.TestAssets;

internal static partial class ManagedTextShapingTestAssets {
    /// <summary>Small font whose kern context precedes and follows its adjusted input glyph.</summary>
    internal static byte[] CreateContextualPositioningFont(bool extension = false, bool nestedPair = false,
        bool recursive = false, bool unsupportedFormat = false, bool invalidCoverageOffset = false) {
        byte[][] lookups = {
            ContextPositioningLookup(8, ContextPositioningSubtable(nestedPair, recursive,
                unsupportedFormat, invalidCoverageOffset), extension, 9),
            ContextPositioningLookup(nestedPair ? 2 : 1,
                nestedPair ? ContextPairPositioningSubtable(true) : ContextSinglePositioningSubtable(), extension, 9),
            ContextPositioningLookup(2, ContextPairPositioningSubtable(false), extension, 8)
        };
        var data = new ContextPositioningBuilder();
        data.U16(1); data.U16(0);
        int scriptListOffset = data.Reserve(), featureListOffset = data.Reserve(), lookupListOffset = data.Reserve();
        data.Offset(scriptListOffset, 0); int scriptList = data.Position;
        data.U16(1); data.Tag("DFLT"); int scriptOffset = data.Reserve();
        data.Offset(scriptOffset, scriptList); int script = data.Position;
        int languageOffset = data.Reserve(); data.U16(0);
        data.Offset(languageOffset, script); data.U16(0); data.U16(ushort.MaxValue);
        data.U16(1); data.U16(0);
        data.Offset(featureListOffset, 0); int featureList = data.Position;
        data.U16(1); data.Tag("kern"); int featureOffset = data.Reserve();
        data.Offset(featureOffset, featureList); data.U16(0);
        data.U16(nestedPair ? 1 : 2); data.U16(0); if (!nestedPair) data.U16(2);
        data.Offset(lookupListOffset, 0); int lookupList = data.Position;
        data.U16(lookups.Length); var offsets = new int[lookups.Length];
        for (int index = 0; index < offsets.Length; index++) offsets[index] = data.Reserve();
        for (int index = 0; index < offsets.Length; index++) {
            data.Offset(offsets[index], lookupList); data.Bytes(lookups[index]);
        }
        // Glyph IDs 1/2/3 are bases A/B/C; glyph 4 is the acute mark ignored by kern.
        byte[] gdef = { 0, 1, 0, 0, 0, 12, 0, 0, 0, 0, 0, 0,
            0, 1, 0, 1, 0, 4, 0, 1, 0, 1, 0, 1, 0, 3 };
        return CreateFontFromCmap(CreateDistinctFormat12Cmap(new[] { (int)'A', (int)'B', (int)'C', 0x0301 }),
            glyphCount: 5, gpos: data.ToArray(), additionalTables: new[] { ("GDEF", gdef) });
    }

    private static byte[] ContextPositioningLookup(int type, byte[] subtable, bool extension, int flags) {
        var data = new ContextPositioningBuilder();
        data.U16(extension ? 9 : type); data.U16(flags); data.U16(1); data.U16(8);
        if (extension) { data.U16(1); data.U16(type); data.U32(8); }
        data.Bytes(subtable); return data.ToArray();
    }

    private static byte[] ContextPositioningSubtable(bool pair, bool recursive, bool unsupported, bool invalidCoverage) {
        var data = new ContextPositioningBuilder();
        data.U16(unsupported ? 2 : 3); data.U16(1); int back = data.Reserve();
        data.U16(pair ? 2 : 1); int first = data.Reserve(), second = pair ? data.Reserve() : -1;
        data.U16(pair ? 0 : 1); int following = pair ? -1 : data.Reserve();
        data.U16(1); data.U16(0); data.U16(recursive ? 0 : 1);
        data.Offset(back, 0); data.Coverage(1);
        if (!invalidCoverage) data.Offset(first, 0); else data.Set(first, ushort.MaxValue);
        data.Coverage(2);
        if (second >= 0) { data.Offset(second, 0); data.Coverage(3); }
        if (following >= 0) { data.Offset(following, 0); data.Coverage(3); }
        return data.ToArray();
    }

    private static byte[] ContextSinglePositioningSubtable() {
        var data = new ContextPositioningBuilder();
        data.U16(1); int coverage = data.Reserve(); data.U16(6);
        data.U16(20); data.U16(-80);
        data.Offset(coverage, 0); data.Coverage(2); return data.ToArray();
    }

    private static byte[] ContextPairPositioningSubtable(bool secondValue) {
        var data = new ContextPositioningBuilder();
        data.U16(1); int coverage = data.Reserve(); data.U16(5); data.U16(secondValue ? 2 : 0);
        data.U16(1); int pair = data.Reserve();
        data.Offset(coverage, 0); data.Coverage(2);
        data.Offset(pair, 0); data.U16(1); data.U16(3); data.U16(7); data.U16(-25);
        if (secondValue) data.U16(30);
        return data.ToArray();
    }

    private sealed class ContextPositioningBuilder {
        private readonly List<byte> _bytes = new();
        internal int Position => _bytes.Count;
        internal void U16(int value) { _bytes.Add(unchecked((byte)(value >> 8))); _bytes.Add(unchecked((byte)value)); }
        internal void U32(int value) { U16(value >> 16); U16(value); }
        internal int Reserve() { int offset = Position; U16(0); return offset; }
        internal void Set(int offset, int value) { _bytes[offset] = (byte)(value >> 8); _bytes[offset + 1] = (byte)value; }
        internal void Offset(int offset, int origin) => Set(offset, checked(Position - origin));
        internal void Bytes(byte[] bytes) => _bytes.AddRange(bytes);
        internal void Tag(string tag) { foreach (char character in tag) _bytes.Add((byte)character); }
        internal void Coverage(int glyph) { U16(1); U16(1); U16(glyph); }
        internal byte[] ToArray() => _bytes.ToArray();
    }
}
