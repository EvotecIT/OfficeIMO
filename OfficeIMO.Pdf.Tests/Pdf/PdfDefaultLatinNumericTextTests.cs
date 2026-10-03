using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfDefaultLatinNumericTextTests {
    [Fact]
    public void NumericAutoDirectionRetainsLogicalIndicesWithoutSelectingUnsupportedLatinFeatures() {
        byte[] data = ManagedTextShapingTestAssets.CreateFontWithCommonSubstitution(true);
        var request = new OfficeTextShapingRequest("11,1", "Numeric", data, false, 1000,
            OfficeTextDirection.Auto, null, default, fontCollectionIndex: null, variationCoordinates: null,
            cloneFontData: true, applyDefaultLatinLigatures: true);
        OfficeTextShapingResult? result = OfficeManagedTextShapingProvider.Instance.ShapeText(request);
        Assert.NotNull(result);
        Assert.Equal(OfficeTextDirection.Auto, result!.Direction);
        Assert.Equal("11,1", string.Concat(result.Glyphs.Select(glyph => glyph.UnicodeText)));
        Assert.Equal(new[] { 0, 1, 2, 3 }, result.Glyphs.Select(glyph => glyph.TextIndex));
        Assert.Equal(new[] { 3, 3, 2, 3 }, result.Glyphs.Select(glyph => glyph.GlyphId));
        Assert.All(Enumerable.Range(0, 4), index => Assert.Equal(0, result.GetAdvanceAdjustment(index)));
        var latin = new OfficeTextShapingRequest("A,1", "Numeric", data, false, 1000,
            OfficeTextDirection.Auto, null, default, fontCollectionIndex: null, variationCoordinates: null,
            cloneFontData: true, applyDefaultLatinLigatures: true);
        Assert.Null(OfficeManagedTextShapingProvider.Instance.ShapeText(latin));
    }
}
