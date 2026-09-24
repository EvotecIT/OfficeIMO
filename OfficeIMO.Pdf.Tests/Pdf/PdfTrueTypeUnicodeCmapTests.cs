using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfTrueTypeUnicodeCmapTests {
    [Fact]
    public void LargeBmpMapUsesFullUnicodeSubtableWithoutTruncatingLength() {
        var mappings = new SortedDictionary<int, int>();
        for (int scalar = 0x20; scalar < 0x20 + 9000; scalar++) mappings.Add(scalar, scalar - 0x1f);

        byte[] cmap = PdfTrueTypeUnicodeCmap.BuildUnicodeCmap(mappings);

        Assert.Equal(1, ReadUInt16(cmap, 2));
        Assert.Equal(3, ReadUInt16(cmap, 4));
        Assert.Equal(10, ReadUInt16(cmap, 6));
        int subtable = checked((int)ReadUInt32(cmap, 8));
        Assert.Equal(12, ReadUInt16(cmap, subtable));
        Assert.Equal((uint)(cmap.Length - subtable), ReadUInt32(cmap, subtable + 4));
        // Consecutive code points with consecutive glyph IDs form one format-12 group.
        Assert.Equal(1U, ReadUInt32(cmap, subtable + 12));
        Assert.Equal(0x20U, ReadUInt32(cmap, subtable + 16));
        Assert.Equal((uint)(0x20 + mappings.Count - 1), ReadUInt32(cmap, subtable + 20));
        Assert.Equal(1U, ReadUInt32(cmap, subtable + 24));
    }

    private static ushort ReadUInt16(byte[] data, int offset) => (ushort)((data[offset] << 8) | data[offset + 1]);
    private static uint ReadUInt32(byte[] data, int offset) =>
        ((uint)data[offset] << 24) | ((uint)data[offset + 1] << 16) | ((uint)data[offset + 2] << 8) | data[offset + 3];
}
