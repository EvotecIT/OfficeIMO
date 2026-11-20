using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed class DrawingColorEmojiGlyphTests {
    // Noto Color Emoji stores PNG glyphs in CBDT/CBLC and Apple Color Emoji in sbix. The fixture PNG is
    // 0.8 em square on the baseline, so a 1000-unit em places it at x 0..800 and y 0..800.
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PngGlyphIsReadWithItsFontUnitBox(bool sbix) {
        OfficeTrueTypeFont font = OfficeTrueTypeFont.TryLoad(ManagedTextShapingTestAssets.CreateBitmapColorFont(0x1F600, sbix))!;

        Assert.True(font.TryGetGlyphMetrics(0x1F600, out int glyph, out _));
        Assert.True(font.TryGetBitmapGlyph(glyph, out OfficeBitmapGlyphImage? image));
        Assert.Equal(new byte[] { 0x89, 0x50, 0x4E, 0x47 }, image!.Png.Take(4));
        Assert.Equal((0D, 0D, 800D, 800D), (image.Left, image.Bottom, image.Width, image.Height));
        Assert.False(font.IsUnpaintableColorGlyph(glyph));
    }

    // Raster and HTML outline routes receive the PNG as a placed box in their y-down output space.
    [Fact]
    public void ColorContoursPlacePngGlyphOnTheBaseline() {
        OfficeTrueTypeFont font = OfficeTrueTypeFont.TryLoad(ManagedTextShapingTestAssets.CreateBitmapColorFont(0x1F600, sbix: false))!;

        Assert.True(font.TryGetColorTextContours("\U0001F600", 10D, 20D, 100D, null, OfficeColor.Black, 10000, default,
            out List<OfficeColorGlyphContours> layers));

        OfficeColorGlyphContours layer = Assert.Single(layers);
        Assert.NotNull(layer.Bitmap);
        double baseline = 20D + 800D / 1000D * 100D;
        Assert.Equal(10D, layer.BitmapLeft, 6);
        Assert.Equal(baseline - 80D, layer.BitmapTop, 6);
        Assert.Equal((80D, 80D), (layer.BitmapWidth, layer.BitmapHeight));
    }

    // Segoe UI Emoji keeps 65,429 entries in one CPAL palette; the reader once rejected every palette over 4,096.
    [Fact]
    public void PaletteLargerThanFourThousandEntriesKeepsItsColorLayers() {
        OfficeTrueTypeFont font = OfficeTrueTypeFont.TryLoad(ManagedTextShapingTestAssets.CreateColorFontWithPaletteSize(0x1F600, 5000))!;

        Assert.True(font.HasColorGlyphs);
        Assert.True(font.TryGetColorTextContours("\U0001F600", 0D, 0D, 100D, null, OfficeColor.Black, 10000, default,
            out List<OfficeColorGlyphContours> layers));
        Assert.Equal(new[] { OfficeColor.FromRgba(255, 0, 0, 255), OfficeColor.FromRgba(0, 0, 255, 255) }, layers.Select(layer => layer.Color));
    }
}
