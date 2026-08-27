using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed class DrawingOpenTypeCmapTests {
    [Fact]
    public void OfficeOpenTypeReaderFallsBackWhenPreferredUnicodeCmapDoesNotMapScalar() {
        byte[] data = ManagedTextShapingTestAssets.CreateFontWithUnicodeCmapFallback('A', 0x1F600);

        OfficeOpenTypeReader reader = Assert.IsType<OfficeOpenTypeReader>(OfficeOpenTypeReader.TryCreate(data));
        OfficeTrueTypeFont font = Assert.IsType<OfficeTrueTypeFont>(OfficeTrueTypeFont.TryLoad(data));

        Assert.Equal(1, reader.MapGlyph('A'));
        Assert.Equal(1, reader.MapGlyph(0x1F600));
        Assert.True(font.HasGlyphs("A" + char.ConvertFromUtf32(0x1F600)));
    }

    [Fact]
    public void UnicodeCmapFallbackUsesTheSameRankedMappingAcrossFontOwners() {
        byte[] data = ManagedTextShapingTestAssets.CreateFontWithConflictingUnicodeCmapFallback('A', 0x1F600);

        OfficeOpenTypeReader reader = Assert.IsType<OfficeOpenTypeReader>(OfficeOpenTypeReader.TryCreate(data));
        OfficeTrueTypeFont font = Assert.IsType<OfficeTrueTypeFont>(OfficeTrueTypeFont.TryLoad(data));

        Assert.Equal(2, reader.MapGlyph('A'));
        Assert.True(font.TryGetGlyphMetrics('A', out int glyphId, out _));
        Assert.Equal(2, glyphId);
    }
}
