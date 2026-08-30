using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed class DrawingOpenTypeKerningTests {
    [Theory]
    [InlineData(-40, 2, -40)]
    [InlineData(0, 2, 0)]
    [InlineData(-40, 3, -80)]
    public void GposPairAdjustmentOverridesLegacyKernWhenPairIsDefined(
        short gposAdjustment,
        ushort gposRightGlyph,
        int expectedAdjustment) {
        byte[] data = CreateKerningTables(
            legacyAdjustment: -80,
            gposAdjustment,
            gposRightGlyph);
        var kerning = new OfficeOpenTypeKerning(
            data,
            kern: 0,
            gpos: 64,
            includeExtendedGpos: true);

        Assert.Equal(expectedAdjustment, kerning.Adjustment(left: 1, right: 2));
    }

    [Fact]
    public void GposKerningUsesOnlyTheActiveScriptsReferencedFeature() {
        byte[] data = CreateKerningTables(
            legacyAdjustment: -80,
            gposAdjustment: -30,
            gposRightGlyph: 2,
            latinAdjustment: -55);
        var kerning = new OfficeOpenTypeKerning(data, kern: 0, gpos: 64, includeExtendedGpos: true);

        Assert.Equal(-30, kerning.Adjustment(left: 1, right: 2, scriptTag: "DFLT"));
        Assert.Equal(-55, kerning.Adjustment(left: 1, right: 2, scriptTag: "latn"));
    }

    [Fact]
    public void LegacyKernUsesHorizontalOrdinarySubtablesAndHonorsOverride() {
        byte[] data = CreateLegacyKernTables(
            (coverage: 1, adjustment: -10),
            (coverage: 0, adjustment: -50),
            (coverage: 5, adjustment: -60),
            (coverage: 9, adjustment: -30),
            (coverage: 1, adjustment: 5));
        var kerning = new OfficeOpenTypeKerning(data, kern: 0, gpos: -1);

        Assert.Equal(-25, kerning.Adjustment(left: 1, right: 2));
    }

    [Fact]
    public void GposPairLookupUsesOnlyTheFirstMatchingSubtable() {
        byte[] data = CreateKerningTables(
            legacyAdjustment: -80,
            gposAdjustment: -30,
            gposRightGlyph: 2);
        const int lookup = 246;
        WriteUInt16(data, lookup + 4, 2);
        WriteUInt16(data, lookup + 6, 10);
        WriteUInt16(data, lookup + 8, 54);
        WritePairSubtable(data, lookup + 10, rightGlyph: 2, adjustment: -30);
        WritePairSubtable(data, lookup + 54, rightGlyph: 2, adjustment: -70);
        var kerning = new OfficeOpenTypeKerning(data, kern: 0, gpos: 64, includeExtendedGpos: true);

        Assert.Equal(-30, kerning.Adjustment(left: 1, right: 2));
    }

    [Fact]
    public void GposPairPositioningAppliesHorizontalFieldsFromBothValueRecords() {
        byte[] data = CreateKerningTables(
            legacyAdjustment: -80,
            gposAdjustment: -30,
            gposRightGlyph: 2);
        WritePairPositioningSubtable(
            data,
            subtable: 254,
            rightGlyph: 2,
            firstXPlacement: -10,
            firstXAdvance: -20,
            secondXPlacement: -30,
            secondXAdvance: -40);
        var kerning = new OfficeOpenTypeKerning(data, kern: 0, gpos: 64, includeExtendedGpos: true);

        OfficeOpenTypePairPositioning positioning = kerning.Positioning(left: 1, right: 2, scriptTag: "DFLT");

        Assert.Equal(-10, positioning.FirstGlyphXPlacement);
        Assert.Equal(-20, positioning.FirstGlyphXAdvance);
        Assert.Equal(-30, positioning.SecondGlyphXPlacement);
        Assert.Equal(-40, positioning.SecondGlyphXAdvance);
        Assert.Equal(-60, kerning.Adjustment(left: 1, right: 2));
    }

    [Fact]
    public void GposClassPairPositioningAppliesHorizontalFieldsFromBothValueRecords() {
        byte[] data = CreateKerningTables(
            legacyAdjustment: -80,
            gposAdjustment: -30,
            gposRightGlyph: 2);
        WriteClassPairPositioningSubtable(
            data,
            subtable: 254,
            firstXPlacement: 11,
            firstXAdvance: 22,
            secondXPlacement: 33,
            secondXAdvance: 44);
        var kerning = new OfficeOpenTypeKerning(data, kern: 0, gpos: 64);

        OfficeOpenTypePairPositioning positioning = kerning.Positioning(left: 1, right: 2, scriptTag: "DFLT");

        Assert.Equal(11, positioning.FirstGlyphXPlacement);
        Assert.Equal(22, positioning.FirstGlyphXAdvance);
        Assert.Equal(33, positioning.SecondGlyphXPlacement);
        Assert.Equal(44, positioning.SecondGlyphXAdvance);
        Assert.Equal(66, kerning.Adjustment(left: 1, right: 2));
    }

    [Fact]
    public void GposPairPositioningAppliesVariationIndexAdjustments() {
        byte[] data = CreateKerningTables(
            legacyAdjustment: -80,
            gposAdjustment: -30,
            gposRightGlyph: 2);
        WritePairVariationPositioningSubtable(data, subtable: 254, rightGlyph: 2);
        var kerning = new OfficeOpenTypeKerning(
            data,
            kern: 0,
            gpos: 64,
            includeExtendedGpos: true,
            variationDelta: (outer, inner) => checked((outer * 100) + inner));

        OfficeOpenTypePairPositioning positioning = kerning.Positioning(left: 1, right: 2, scriptTag: "DFLT");

        Assert.Equal(193, positioning.FirstGlyphXPlacement);
        Assert.Equal(385, positioning.FirstGlyphXAdvance);
    }

    [Fact]
    public void GposClassPairPositioningAppliesVariationIndexAdjustments() {
        byte[] data = CreateKerningTables(
            legacyAdjustment: -80,
            gposAdjustment: -30,
            gposRightGlyph: 2);
        WriteClassPairVariationPositioningSubtable(data, subtable: 254);
        var kerning = new OfficeOpenTypeKerning(
            data,
            kern: 0,
            gpos: 64,
            variationDelta: (outer, inner) => checked((outer * 100) + inner));

        OfficeOpenTypePairPositioning positioning = kerning.Positioning(left: 1, right: 2, scriptTag: "DFLT");

        Assert.Equal(193, positioning.FirstGlyphXPlacement);
        Assert.Equal(385, positioning.FirstGlyphXAdvance);
    }

    [Fact]
    public void GposRunSkipsTheSecondGlyphWhenValueRecord2IsPresent() {
        byte[] data = CreateKerningTables(
            legacyAdjustment: 0,
            gposAdjustment: 0,
            gposRightGlyph: 2);
        WriteSequencedPairPositioningSubtable(data, subtable: 254);
        var kerning = new OfficeOpenTypeKerning(data, kern: 0, gpos: 64);

        OfficeOpenTypeGlyphPositioning[] positioning = kerning.PositionRun(
            new[] { 1, 2, 3 },
            new[] { 0, 0, 0 });

        Assert.Equal(-10, positioning[0].XAdvance);
        Assert.Equal(-20, positioning[1].XAdvance);
        Assert.Equal(0, positioning[2].XAdvance);
    }

    private static byte[] CreateKerningTables(
        short legacyAdjustment,
        short gposAdjustment,
        ushort gposRightGlyph,
        short latinAdjustment = -20) {
        var data = new byte[384];

        WriteUInt16(data, 2, 1);       // kern subtable count
        WriteUInt16(data, 6, 20);      // kern subtable length
        WriteUInt16(data, 8, 1);       // horizontal format 0
        WriteUInt16(data, 10, 1);      // pair count
        WriteUInt16(data, 18, 1);      // left glyph
        WriteUInt16(data, 20, 2);      // right glyph
        WriteInt16(data, 22, legacyAdjustment);

        WriteUInt16(data, 64, 1);      // GPOS major version
        WriteUInt16(data, 68, 16);     // ScriptList offset
        WriteUInt16(data, 70, 96);     // FeatureList offset
        WriteUInt16(data, 72, 176);    // LookupList offset

        WriteUInt16(data, 80, 2);      // script count
        WriteTag(data, 82, "DFLT");
        WriteUInt16(data, 86, 14);     // default script table
        WriteTag(data, 88, "latn");
        WriteUInt16(data, 92, 28);     // Latin script table
        WriteUInt16(data, 94, 4);      // default LangSys offset
        WriteUInt16(data, 98, 0);      // lookup order
        WriteUInt16(data, 100, ushort.MaxValue);
        WriteUInt16(data, 102, 1);     // feature count
        WriteUInt16(data, 104, 0);     // default feature
        WriteUInt16(data, 108, 4);     // Latin default LangSys offset
        WriteUInt16(data, 112, 0);
        WriteUInt16(data, 114, ushort.MaxValue);
        WriteUInt16(data, 116, 1);
        WriteUInt16(data, 118, 1);     // Latin feature

        WriteUInt16(data, 160, 2);     // feature count
        WriteTag(data, 162, "kern");
        WriteUInt16(data, 166, 16);
        WriteTag(data, 168, "kern");
        WriteUInt16(data, 172, 22);
        WriteUInt16(data, 176, 0);
        WriteUInt16(data, 178, 1);
        WriteUInt16(data, 180, 0);     // default lookup
        WriteUInt16(data, 182, 0);
        WriteUInt16(data, 184, 1);
        WriteUInt16(data, 186, 1);     // Latin lookup

        WriteUInt16(data, 240, 2);     // lookup count
        WriteUInt16(data, 242, 6);
        WriteUInt16(data, 244, 40);
        WritePairLookup(data, lookupOffset: 246, gposRightGlyph, gposAdjustment);
        WritePairLookup(data, lookupOffset: 280, rightGlyph: 2, latinAdjustment);
        return data;
    }

    private static byte[] CreateLegacyKernTables(params (ushort coverage, short adjustment)[] tables) {
        var data = new byte[4 + (tables.Length * 20)];
        WriteUInt16(data, 2, checked((ushort)tables.Length));
        for (int index = 0; index < tables.Length; index++) {
            int offset = 4 + (index * 20);
            WriteUInt16(data, offset + 2, 20);
            WriteUInt16(data, offset + 4, tables[index].coverage);
            WriteUInt16(data, offset + 6, 1);
            WriteUInt16(data, offset + 14, 1);
            WriteUInt16(data, offset + 16, 2);
            WriteInt16(data, offset + 18, tables[index].adjustment);
        }
        return data;
    }

    private static void WritePairLookup(byte[] data, int lookupOffset, ushort rightGlyph, short adjustment) {
        WriteUInt16(data, lookupOffset, 2);
        WriteUInt16(data, lookupOffset + 4, 1);
        WriteUInt16(data, lookupOffset + 6, 8);
        WritePairSubtable(data, lookupOffset + 8, rightGlyph, adjustment);
    }

    private static void WritePairSubtable(byte[] data, int subtable, ushort rightGlyph, short adjustment) {
        WriteUInt16(data, subtable, 1);
        WriteUInt16(data, subtable + 2, 12);
        WriteUInt16(data, subtable + 4, 4);
        WriteUInt16(data, subtable + 8, 1);
        WriteUInt16(data, subtable + 10, 18);
        WriteUInt16(data, subtable + 12, 1);
        WriteUInt16(data, subtable + 14, 1);
        WriteUInt16(data, subtable + 16, 1);
        WriteUInt16(data, subtable + 18, 1);
        WriteUInt16(data, subtable + 20, rightGlyph);
        WriteInt16(data, subtable + 22, adjustment);
    }

    private static void WritePairPositioningSubtable(
        byte[] data,
        int subtable,
        ushort rightGlyph,
        short firstXPlacement,
        short firstXAdvance,
        short secondXPlacement,
        short secondXAdvance) {
        WriteUInt16(data, subtable, 1);
        WriteUInt16(data, subtable + 2, 12);
        WriteUInt16(data, subtable + 4, 5);
        WriteUInt16(data, subtable + 6, 5);
        WriteUInt16(data, subtable + 8, 1);
        WriteUInt16(data, subtable + 10, 18);
        WriteUInt16(data, subtable + 12, 1);
        WriteUInt16(data, subtable + 14, 1);
        WriteUInt16(data, subtable + 16, 1);
        WriteUInt16(data, subtable + 18, 1);
        WriteUInt16(data, subtable + 20, rightGlyph);
        WriteInt16(data, subtable + 22, firstXPlacement);
        WriteInt16(data, subtable + 24, firstXAdvance);
        WriteInt16(data, subtable + 26, secondXPlacement);
        WriteInt16(data, subtable + 28, secondXAdvance);
    }

    private static void WriteClassPairPositioningSubtable(
        byte[] data,
        int subtable,
        short firstXPlacement,
        short firstXAdvance,
        short secondXPlacement,
        short secondXAdvance) {
        WriteUInt16(data, subtable, 2);
        WriteUInt16(data, subtable + 2, 48);
        WriteUInt16(data, subtable + 4, 5);
        WriteUInt16(data, subtable + 6, 5);
        WriteUInt16(data, subtable + 8, 54);
        WriteUInt16(data, subtable + 10, 62);
        WriteUInt16(data, subtable + 12, 2);
        WriteUInt16(data, subtable + 14, 2);

        int classRecord = subtable + 16 + (3 * 8);
        WriteInt16(data, classRecord, firstXPlacement);
        WriteInt16(data, classRecord + 2, firstXAdvance);
        WriteInt16(data, classRecord + 4, secondXPlacement);
        WriteInt16(data, classRecord + 6, secondXAdvance);

        int coverage = subtable + 48;
        WriteUInt16(data, coverage, 1);
        WriteUInt16(data, coverage + 2, 1);
        WriteUInt16(data, coverage + 4, 1);

        int classDef1 = subtable + 54;
        WriteUInt16(data, classDef1, 1);
        WriteUInt16(data, classDef1 + 2, 1);
        WriteUInt16(data, classDef1 + 4, 1);
        WriteUInt16(data, classDef1 + 6, 1);

        int classDef2 = subtable + 62;
        WriteUInt16(data, classDef2, 1);
        WriteUInt16(data, classDef2 + 2, 2);
        WriteUInt16(data, classDef2 + 4, 1);
        WriteUInt16(data, classDef2 + 6, 1);
    }

    private static void WritePairVariationPositioningSubtable(byte[] data, int subtable, ushort rightGlyph) {
        WriteUInt16(data, subtable, 1);
        WriteUInt16(data, subtable + 2, 12);
        WriteUInt16(data, subtable + 4, 0x0055);
        WriteUInt16(data, subtable + 6, 0);
        WriteUInt16(data, subtable + 8, 1);
        WriteUInt16(data, subtable + 10, 18);
        WriteUInt16(data, subtable + 12, 1);
        WriteUInt16(data, subtable + 14, 1);
        WriteUInt16(data, subtable + 16, 1);
        WriteUInt16(data, subtable + 18, 1);
        WriteUInt16(data, subtable + 20, rightGlyph);
        WriteInt16(data, subtable + 22, -10);
        WriteInt16(data, subtable + 24, -20);
        WriteUInt16(data, subtable + 26, 46);
        WriteUInt16(data, subtable + 28, 52);
        WriteVariationIndex(data, subtable + 64, outer: 2, inner: 3);
        WriteVariationIndex(data, subtable + 70, outer: 4, inner: 5);
    }

    private static void WriteClassPairVariationPositioningSubtable(byte[] data, int subtable) {
        WriteUInt16(data, subtable, 2);
        WriteUInt16(data, subtable + 2, 64);
        WriteUInt16(data, subtable + 4, 0x0055);
        WriteUInt16(data, subtable + 6, 0);
        WriteUInt16(data, subtable + 8, 70);
        WriteUInt16(data, subtable + 10, 78);
        WriteUInt16(data, subtable + 12, 2);
        WriteUInt16(data, subtable + 14, 2);

        int classRecord = subtable + 16 + (3 * 8);
        WriteInt16(data, classRecord, -10);
        WriteInt16(data, classRecord + 2, -20);
        WriteUInt16(data, classRecord + 4, 52);
        WriteUInt16(data, classRecord + 6, 58);
        WriteVariationIndex(data, subtable + 52, outer: 2, inner: 3);
        WriteVariationIndex(data, subtable + 58, outer: 4, inner: 5);

        int coverage = subtable + 64;
        WriteUInt16(data, coverage, 1);
        WriteUInt16(data, coverage + 2, 1);
        WriteUInt16(data, coverage + 4, 1);

        int classDef1 = subtable + 70;
        WriteUInt16(data, classDef1, 1);
        WriteUInt16(data, classDef1 + 2, 1);
        WriteUInt16(data, classDef1 + 4, 1);
        WriteUInt16(data, classDef1 + 6, 1);

        int classDef2 = subtable + 78;
        WriteUInt16(data, classDef2, 1);
        WriteUInt16(data, classDef2 + 2, 2);
        WriteUInt16(data, classDef2 + 4, 1);
        WriteUInt16(data, classDef2 + 6, 1);
    }

    private static void WriteVariationIndex(byte[] data, int offset, ushort outer, ushort inner) {
        WriteUInt16(data, offset, outer);
        WriteUInt16(data, offset + 2, inner);
        WriteUInt16(data, offset + 4, 0x8000);
    }

    private static void WriteSequencedPairPositioningSubtable(byte[] data, int subtable) {
        WriteUInt16(data, subtable, 1);
        WriteUInt16(data, subtable + 2, 14);
        WriteUInt16(data, subtable + 4, 4);
        WriteUInt16(data, subtable + 6, 4);
        WriteUInt16(data, subtable + 8, 2);
        WriteUInt16(data, subtable + 10, 22);
        WriteUInt16(data, subtable + 12, 34);

        int coverage = subtable + 14;
        WriteUInt16(data, coverage, 1);
        WriteUInt16(data, coverage + 2, 2);
        WriteUInt16(data, coverage + 4, 1);
        WriteUInt16(data, coverage + 6, 2);

        int firstPairSet = subtable + 22;
        WriteUInt16(data, firstPairSet, 1);
        WriteUInt16(data, firstPairSet + 2, 2);
        WriteInt16(data, firstPairSet + 4, -10);
        WriteInt16(data, firstPairSet + 6, -20);

        int secondPairSet = subtable + 34;
        WriteUInt16(data, secondPairSet, 1);
        WriteUInt16(data, secondPairSet + 2, 3);
        WriteInt16(data, secondPairSet + 4, -30);
        WriteInt16(data, secondPairSet + 6, -40);
    }

    private static void WriteTag(byte[] data, int offset, string tag) {
        for (int index = 0; index < 4; index++) data[offset + index] = (byte)tag[index];
    }

    private static void WriteUInt16(byte[] data, int offset, ushort value) {
        data[offset] = (byte)(value >> 8);
        data[offset + 1] = (byte)value;
    }

    private static void WriteInt16(byte[] data, int offset, short value) =>
        WriteUInt16(data, offset, unchecked((ushort)value));
}
