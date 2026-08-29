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
