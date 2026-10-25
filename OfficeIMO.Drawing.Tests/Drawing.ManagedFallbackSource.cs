using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed class DrawingManagedFallbackSourceTests {
    [Theory]
    [InlineData("\u0430\u0483")]
    [InlineData("\u0430\uA670")]
    [InlineData("\U0001D15F\U0001D165")]
    [InlineData("\u0627\u064E")]
    [InlineData("\u05D0\u05B7")]
    public void DiagnosedMarkFallbackRetainsOriginalScalarsWhenEveryCmapEntryAliasesOneGlyph(string source) {
        byte[] data = ManagedTextShapingTestAssets.CreateFont(0x0627, 0xFE8D, 0x064E, 0x05D0, 0x05B7,
            0x0430, 0x0483, 0x1D15F, 0x1D165, 0xA670);
        var request = new OfficeTextShapingRequest(source, "Aliased fallback", data, false, 1000,
            OfficeTextElements.ResolveBaseDirection(source));

        // Cmap coverage does not supply the missing GPOS anchors. Complete
        // admission remains declined; only the diagnosed approximation is kept.
        Assert.Null(OfficeManagedTextShapingProvider.Instance.ShapeText(request));
        OfficeTextShapingResult fallback = Assert.IsType<OfficeTextShapingResult>(
            OfficeManagedTextShapingProvider.Instance.ShapeDiagnosedFallback(request));

        Assert.Equal(2, fallback.Glyphs.Count);
        Assert.All(fallback.Glyphs, glyph => {
            Assert.Equal(1, glyph.GlyphId);
            Assert.Equal(glyph.UnicodeText, source.Substring(glyph.TextIndex, glyph.UnicodeText.Length));
        });
        Assert.Equal(source, string.Concat(fallback.Glyphs.OrderBy(glyph => glyph.TextIndex).Select(glyph => glyph.UnicodeText)));
    }
}
