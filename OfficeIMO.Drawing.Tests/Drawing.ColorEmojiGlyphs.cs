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

    // A duplicate record shares its target's artwork instead of copying the PNG again.
    [Fact]
    public void SbixDuplicateSharesItsTargetsPng() {
        OfficeTrueTypeFont font = OfficeTrueTypeFont.TryLoad(ManagedTextShapingTestAssets.CreateSbixVariantFont(ManagedTextShapingTestAssets.SbixVariant.Duplicate))!;
        Assert.True(font.TryGetGlyphMetrics(0x1F601, out int duplicate, out _));
        Assert.True(font.TryGetGlyphMetrics(0x1F600, out int target, out _));

        Assert.True(font.TryGetBitmapGlyph(duplicate, out OfficeBitmapGlyphImage? shared));
        Assert.True(font.TryGetBitmapGlyph(target, out OfficeBitmapGlyphImage? original));

        Assert.Same(original, shared);
    }

    // Strikes need not cover every glyph: a glyph missing from the largest strike comes from the next one.
    [Fact]
    public void SbixGlyphMissingFromTheLargestStrikeComesFromASmallerOne() {
        OfficeTrueTypeFont font = OfficeTrueTypeFont.TryLoad(ManagedTextShapingTestAssets.CreateSbixVariantFont(ManagedTextShapingTestAssets.SbixVariant.MissingFromLargestStrike))!;
        Assert.True(font.TryGetGlyphMetrics(0x1F600, out int glyph, out _));

        Assert.True(font.TryGetBitmapGlyph(glyph, out OfficeBitmapGlyphImage? image));
        Assert.Equal((800D, 800D), (image!.Width, image.Height));
    }

    // sbix origin offsets are measured from the outline's bounding-box origin when the glyph has an outline.
    [Fact]
    public void SbixBitmapOfAnOutlinedGlyphStartsAtTheOutlineOrigin() {
        OfficeTrueTypeFont font = OfficeTrueTypeFont.TryLoad(ManagedTextShapingTestAssets.CreateSbixVariantFont(ManagedTextShapingTestAssets.SbixVariant.OutlineOrigin))!;
        Assert.True(font.TryGetGlyphMetrics(0x1F600, out int glyph, out _));

        Assert.True(font.TryGetBitmapGlyph(glyph, out OfficeBitmapGlyphImage? image));
        Assert.Equal((ManagedTextShapingTestAssets.SbixOutlineOrigin.Left, ManagedTextShapingTestAssets.SbixOutlineOrigin.Bottom),
            ((int)image!.Left, (int)image.Bottom));
    }

    // Raster text draws PNG color glyphs, as the HTML and PDF routes do.
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RasterCanvasDrawsPngColorGlyphs(bool sbix) {
        var fonts = new OfficeFontFaceCollection().Add("Bitmap emoji", ManagedTextShapingTestAssets.CreateBitmapColorFont(0x1F600, sbix));
        var image = new OfficeRasterImage(300, 150, OfficeColor.White);
        var canvas = new OfficeRasterCanvas(image, null, fonts);

        canvas.DrawText("\U0001F600\U0001F600", 10D, 10D, 280D, 130D, OfficeColor.Black, 60D, fontFamily: "Bitmap emoji");

        int red = 0, left = int.MaxValue, right = -1;
        for (int y = 0; y < image.Height; y++) {
            for (int x = 0; x < image.Width; x++) {
                OfficeColor color = image.GetPixel(x, y);
                if (color.R <= 200 || color.G >= 60 || color.B >= 60) continue;
                red++; left = Math.Min(left, x); right = Math.Max(right, x);
            }
        }
        Assert.True(red > 0, "The PNG glyphs are drawn.");
        // Two 48 px squares (0.8 em at 60 px) one 30 px advance apart: 78 px wide, their union 3744 px.
        Assert.InRange(red, 3744 * 0.93D, 3744 * 1.07D);
        Assert.InRange(right - left + 1, 76, 80);
    }
}
