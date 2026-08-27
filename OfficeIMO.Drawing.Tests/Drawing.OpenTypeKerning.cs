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

    private static byte[] CreateKerningTables(
        short legacyAdjustment,
        short gposAdjustment,
        ushort gposRightGlyph) {
        var data = new byte[128];

        WriteUInt16(data, 2, 1);       // kern subtable count
        WriteUInt16(data, 6, 20);      // kern subtable length
        WriteUInt16(data, 8, 1);       // horizontal format 0
        WriteUInt16(data, 10, 1);      // pair count
        WriteUInt16(data, 18, 1);      // left glyph
        WriteUInt16(data, 20, 2);      // right glyph
        WriteInt16(data, 22, legacyAdjustment);

        WriteUInt16(data, 64, 1);      // GPOS major version
        WriteUInt16(data, 70, 10);     // feature list offset
        WriteUInt16(data, 72, 24);     // lookup list offset
        WriteUInt16(data, 74, 1);      // feature count
        data[76] = (byte)'k';
        data[77] = (byte)'e';
        data[78] = (byte)'r';
        data[79] = (byte)'n';
        WriteUInt16(data, 80, 8);      // feature table offset
        WriteUInt16(data, 84, 1);      // feature lookup count
        WriteUInt16(data, 88, 1);      // lookup count
        WriteUInt16(data, 90, 4);      // lookup table offset
        WriteUInt16(data, 92, 2);      // pair-adjustment lookup
        WriteUInt16(data, 96, 1);      // subtable count
        WriteUInt16(data, 98, 8);      // subtable offset
        WriteUInt16(data, 100, 1);     // PairPos format 1
        WriteUInt16(data, 102, 12);    // coverage offset
        WriteUInt16(data, 104, 4);     // first value: xAdvance
        WriteUInt16(data, 108, 1);     // pair-set count
        WriteUInt16(data, 110, 18);    // pair-set offset
        WriteUInt16(data, 112, 1);     // coverage format 1
        WriteUInt16(data, 114, 1);     // covered glyph count
        WriteUInt16(data, 116, 1);     // covered left glyph
        WriteUInt16(data, 118, 1);     // pair value count
        WriteUInt16(data, 120, gposRightGlyph);
        WriteInt16(data, 122, gposAdjustment);
        return data;
    }

    private static void WriteUInt16(byte[] data, int offset, ushort value) {
        data[offset] = (byte)(value >> 8);
        data[offset + 1] = (byte)value;
    }

    private static void WriteInt16(byte[] data, int offset, short value) =>
        WriteUInt16(data, offset, unchecked((ushort)value));
}
