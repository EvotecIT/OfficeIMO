using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Theory]
    [InlineData(2048)]
    [InlineData(8189)]
    public void OpenTypeReadersAcceptLargeValidFormat4Cmaps(int segmentCount) {
        // CJK fonts can exceed 1,024 sparse segments. Use the format's byte-length
        // bound rather than requiring an installed proprietary font on CI hosts.
        int length = 16 + segmentCount * 8;
        var cmap = new byte[12 + length];
        WriteUInt16(cmap, 2, 1);
        WriteUInt16(cmap, 4, 3);
        WriteUInt16(cmap, 6, 1);
        WriteUInt32(cmap, 8, 12);
        WriteUInt16(cmap, 12, 4);
        WriteUInt16(cmap, 14, length);
        WriteUInt16(cmap, 18, segmentCount * 2);
        int power = 1;
        int selector = 0;
        while (power * 2 <= segmentCount) { power *= 2; selector++; }
        WriteUInt16(cmap, 20, power * 2);
        WriteUInt16(cmap, 22, selector);
        WriteUInt16(cmap, 24, segmentCount * 2 - power * 2);
        int ends = 26;
        int starts = ends + segmentCount * 2 + 2;
        int deltas = starts + segmentCount * 2;
        for (int index = 0; index < segmentCount; index++) {
            int scalar = index == segmentCount - 1 ? 0xFFFF : 32 + index * 2;
            WriteUInt16(cmap, ends + index * 2, scalar);
            WriteUInt16(cmap, starts + index * 2, scalar);
            WriteUInt16(cmap, deltas + index * 2,
                index == segmentCount - 1 ? 1 : unchecked((ushort)(1 - scalar)));
        }

        byte[] data = CreateMinimalTrueTypeFont(cmap);
        var reader = Assert.IsType<OfficeOpenTypeReader>(OfficeOpenTypeReader.TryCreate(data));
        var font = Assert.IsType<OfficeTrueTypeFont>(OfficeTrueTypeFont.TryLoad(data));
        foreach (int index in new[] { 0, segmentCount / 2, segmentCount - 2 }) {
            char scalar = (char)(32 + index * 2);
            Assert.Equal(1, reader.MapGlyph(scalar));
            Assert.True(font.HasGlyphs(scalar.ToString()));
            Assert.False(font.HasGlyphs(((char)(scalar + 1)).ToString()));
        }
        Assert.Equal(0, reader.MapGlyph(0xFFFF));
    }
}
