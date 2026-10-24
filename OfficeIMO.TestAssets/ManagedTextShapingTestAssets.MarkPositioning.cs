using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.TestAssets;

internal static partial class ManagedTextShapingTestAssets {
    /// <summary>Design-unit anchor fixture with a base, three marks and a two-component ligature.</summary>
    internal static byte[] CreateMarkPositioningFont(bool extension = false, bool unsupportedAnchorDevice = false,
        bool missingBaseAnchor = false, bool filteredMarkStack = true, bool extraLigatureComponent = false,
        ushort baseLookupFlags = 0, bool classifyBaseAsLigature = false, bool missingParentCoverage = false) {
        byte[] baseAttachment = AttachmentSubtable(4, unsupportedAnchorDevice, missingBaseAnchor,
            missingParentCoverage: missingParentCoverage);
        byte[] ligatureAttachment = AttachmentSubtable(5, false, false, extraLigatureComponent);
        byte[] markAttachment = AttachmentSubtable(6, false, false);
        byte[][] lookups = {
            AttachmentLookup(4, baseAttachment, extension, baseLookupFlags),
            AttachmentLookup(5, ligatureAttachment, extension, 0),
            AttachmentLookup(6, markAttachment, extension, filteredMarkStack ? (ushort)16 : (ushort)0)
        };
        var gpos = new MarkTableBuilder();
        gpos.U16(1); gpos.U16(0); int scriptsOffset = gpos.Reserve(), featuresOffset = gpos.Reserve(), lookupsOffset = gpos.Reserve();
        gpos.Offset(scriptsOffset, 0);
        int scriptList = gpos.Position; gpos.U16(1); gpos.Tag("DFLT"); int scriptOffset = gpos.Reserve();
        gpos.Offset(scriptOffset, scriptList); int script = gpos.Position; int langOffset = gpos.Reserve(); gpos.U16(0);
        gpos.Offset(langOffset, script); gpos.U16(0); gpos.U16(ushort.MaxValue); gpos.U16(2); gpos.U16(0); gpos.U16(1);
        gpos.Offset(featuresOffset, 0);
        int featureList = gpos.Position; gpos.U16(2); gpos.Tag("mark"); int markOffset = gpos.Reserve(); gpos.Tag("mkmk"); int stackOffset = gpos.Reserve();
        gpos.Offset(markOffset, featureList); gpos.U16(0); gpos.U16(2); gpos.U16(0); gpos.U16(1);
        gpos.Offset(stackOffset, featureList); gpos.U16(0); gpos.U16(1); gpos.U16(2);
        gpos.Offset(lookupsOffset, 0);
        int lookupList = gpos.Position; gpos.U16(lookups.Length); int[] lookupOffsets = lookups.Select(_ => gpos.Reserve()).ToArray();
        for (int index = 0; index < lookups.Length; index++) { gpos.Offset(lookupOffsets[index], lookupList); gpos.Bytes(lookups[index]); }
        return CreateFontFromCmap(CreateDistinctFormat12Cmap(new[] { (int)'A', 0x0301, 0x0302, 0x0303 }), glyphCount: 6,
            gpos: gpos.ToArray(), additionalTables: new[] { ("GDEF", MarkFilterGdef(classifyBaseAsLigature)) });
    }

    private static byte[] AttachmentLookup(ushort type, byte[] subtable, bool extension, ushort flags) {
        var data = new MarkTableBuilder(); data.U16(extension ? 9 : type); data.U16(flags); data.U16(1);
        int subtableOffset = data.Reserve(); if ((flags & 16) != 0) data.U16(0);
        data.Offset(subtableOffset, 0);
        if (extension) { data.U16(1); data.U16(type); data.U32(8); }
        data.Bytes(subtable); return data.ToArray();
    }

    private static byte[] AttachmentSubtable(int type, bool unsupportedDevice, bool missingBaseAnchor,
        bool extraLigatureComponent = false, bool missingParentCoverage = false) {
        var data = new MarkTableBuilder(); data.U16(1);
        int markCoverageOffset = data.Reserve(), parentCoverageOffset = data.Reserve(); data.U16(1);
        int markArrayOffset = data.Reserve(), parentArrayOffset = data.Reserve();
        data.Offset(markCoverageOffset, 0); data.Coverage(type == 6 ? new[] { 4 } : new[] { 2, 3, 4 });
        data.Offset(parentCoverageOffset, 0); data.Coverage(new[] { type == 5 || missingParentCoverage ? 5 : type == 6 ? 2 : 1 });
        data.Offset(markArrayOffset, 0);
        int markArray = data.Position, count = type == 6 ? 1 : 3; data.U16(count);
        var markAnchors = new List<int>();
        for (int index = 0; index < count; index++) { data.U16(0); markAnchors.Add(data.Reserve()); }
        foreach (int offset in markAnchors) {
            data.Offset(offset, markArray); data.U16(unsupportedDevice ? 3 : 1); data.U16(type == 6 ? 1 : 10); data.U16(type == 6 ? 2 : 20);
            if (unsupportedDevice) { data.U16(1); data.U16(0); }
        }
        data.Offset(parentArrayOffset, 0);
        int parentArray = data.Position; data.U16(1); int parentOffset = data.Reserve();
        if (type == 5) {
            data.Offset(parentOffset, parentArray); int ligature = data.Position; data.U16(extraLigatureComponent ? 3 : 2);
            int first = data.Reserve(), second = data.Reserve();
            int third = extraLigatureComponent ? data.Reserve() : -1;
            data.Offset(first, ligature); data.Anchor(100, 600);
            data.Offset(second, ligature); data.Anchor(300, 500);
            if (third >= 0) { data.Offset(third, ligature); data.Anchor(500, 400); }
        } else if (!missingBaseAnchor) {
            data.Offset(parentOffset, parentArray); data.Anchor(type == 6 ? 40 : 200, type == 6 ? 80 : 700);
        }
        return data.ToArray();
    }

    private static byte[] MarkFilterGdef(bool classifyBaseAsLigature) {
        var data = new MarkTableBuilder(); data.U16(1); data.U16(2);
        int classes = data.Reserve(); data.U16(0); data.U16(0); data.U16(0); int sets = data.Reserve();
        data.Offset(classes, 0); data.U16(1); data.U16(1); data.U16(5);
        foreach (int value in new[] { classifyBaseAsLigature ? 2 : 1, 3, 3, 3, 2 }) data.U16(value);
        data.Offset(sets, 0); data.U16(1); data.U16(1); data.U32(8); data.Coverage(new[] { 2, 4 });
        return data.ToArray();
    }

    private sealed class MarkTableBuilder {
        private readonly List<byte> _bytes = new();
        internal int Position => _bytes.Count;
        internal void U16(int value) { _bytes.Add((byte)(value >> 8)); _bytes.Add((byte)value); }
        internal void U32(int value) { U16(value >> 16); U16(value); }
        internal int Reserve() { int offset = Position; U16(0); return offset; }
        internal void Offset(int offset, int origin) { int value = checked(Position - origin); _bytes[offset] = checked((byte)(value >> 8)); _bytes[offset + 1] = (byte)value; }
        internal void Tag(string tag) { foreach (char character in tag) _bytes.Add((byte)character); }
        internal void Bytes(byte[] bytes) => _bytes.AddRange(bytes);
        internal void Coverage(int[] glyphs) { U16(1); U16(glyphs.Length); foreach (int glyph in glyphs) U16(glyph); }
        internal void Anchor(int x, int y) { U16(1); U16(x); U16(y); }
        internal byte[] ToArray() => _bytes.ToArray();
    }
}
