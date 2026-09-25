using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfTrueTypeUnicodeCmapTests {
    [Fact]
    public void LargeSparseMapKeepsEveryGlyphInOneFormat12Subtable() {
        var mappings = new SortedDictionary<int, int>();
        for (int index = 0; index < 9000; index++) mappings.Add(0x1000 + index * 2, index % 2 + 1);

        byte[] cmap = PdfTrueTypeUnicodeCmap.BuildUnicodeCmap(mappings);

        Assert.Equal(1, ReadUInt16(cmap, 2));
        Assert.Equal(3, ReadUInt16(cmap, 4));
        Assert.Equal(10, ReadUInt16(cmap, 6));
        int subtable = checked((int)ReadUInt32(cmap, 8));
        Assert.Equal(12, ReadUInt16(cmap, subtable));
        Assert.Equal(9000U, ReadUInt32(cmap, subtable + 12));
        int lastGroup = subtable + 16 + 8999 * 12;
        Assert.Equal((uint)(0x1000 + 8999 * 2), ReadUInt32(cmap, lastGroup));
        Assert.Equal(2U, ReadUInt32(cmap, lastGroup + 8));

        byte[] source = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs('A', 'B');
        var drawing = new PdfDrawingFontProgram(source, new SortedDictionary<int, int> { ['A'] = 1 },
            _ => 1, _ => false);
        byte[] rebuilt = Assert.IsType<byte[]>(PdfTrueTypeUnicodeCmap.TryAddMappings(drawing, mappings));
        Assert.True(OfficeIMO.Drawing.OfficeTrueTypeFont.TryLoad(rebuilt)?.HasGlyphs(
            char.ConvertFromUtf32(0x1000 + 8999 * 2)));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void UnicodeOnlySimpleFontUsesEncodingDifferenceForItsPaintedGlyph(bool format4) {
        byte[] source = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs('A', 'B');
        if (format4) {
            var original = new PdfDrawingFontProgram(source, new SortedDictionary<int, int> { ['A'] = 1 },
                _ => 1, _ => false);
            source = Assert.IsType<byte[]>(PdfTrueTypeUnicodeCmap.TryAddMappings(original,
                new Dictionary<int, int> { ['B'] = 2 }));
        }
        Assert.True(ToUnicodeCMap.TryParse(Encoding.ASCII.GetBytes(
            "begincmap\n1 beginbfchar\n<41> <0051>\nendbfchar\nendcmap"), out ToUnicodeCMap? cmap));
        var font = new PdfFontResource("F1", "Subset", "WinAnsiEncoding", true, cmap,
            new Dictionary<int, string> { [65] = "B" }, fontSubtype: "TrueType",
            embeddedProgramSubtype: "TrueType", fontDescriptorFlags: 32);

        PdfDrawingFontProgram drawing = Assert.IsType<PdfDrawingFontProgram>(
            PdfTrueTypeUnicodeCmap.TryCreate(font, source, null));

        Assert.Equal(2, drawing.GlyphForCode(65));
        Assert.Equal(2, drawing.UnicodeGlyphs['Q']);
        Assert.Equal(2, drawing.UnicodeGlyphs['B']);
        Assert.True(OfficeIMO.Drawing.OfficeTrueTypeFont.TryLoad(drawing.Program)?.HasGlyphs("Q"));
    }

    [Theory]
    [InlineData("fi", 1)]
    [InlineData("fl", 2)]
    public void LigatureDifferenceLooksUpPaintedGlyphWithoutChangingLogicalCluster(string name, int glyph) {
        byte[] source = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs('\uFB01', '\uFB02');
        var font = new PdfFontResource("F1", "LigatureSubset", "WinAnsiEncoding", false,
            differences: new Dictionary<int, string> { [65] = name }, fontSubtype: "TrueType",
            embeddedProgramSubtype: "TrueType", fontDescriptorFlags: 32);

        PdfDrawingFontProgram drawing = Assert.IsType<PdfDrawingFontProgram>(
            PdfTrueTypeUnicodeCmap.TryCreate(font, source, null));

        Assert.Equal(glyph, drawing.GlyphForCode(65));
        Assert.Equal(name, ResourceResolver.CreateSimpleEncodingDecoder(font)(65));
    }

    [Fact]
    public void GlyfOpenTypeContainerUsesSimpleFontCodeMapping() {
        byte[] source = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs('A', 'B');
        var font = new PdfFontResource("F1", "Subset", "WinAnsiEncoding", false,
            differences: new Dictionary<int, string> { [65] = "B" }, fontSubtype: "TrueType",
            embeddedProgramSubtype: "OpenType", fontDescriptorFlags: 32);

        PdfDrawingFontProgram drawing = Assert.IsType<PdfDrawingFontProgram>(
            PdfTrueTypeUnicodeCmap.TryCreate(font, source, null));

        Assert.Equal(2, drawing.GlyphForCode(65));
    }

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
