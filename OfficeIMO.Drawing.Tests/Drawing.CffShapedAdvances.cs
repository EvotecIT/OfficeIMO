using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingCffShapedAdvanceTests {
    [Theory]
    [InlineData(OfficeTextDirection.LeftToRight)]
    [InlineData(OfficeTextDirection.RightToLeft)]
    public void CffMeasurementAndContoursRetainManagedDefaultPairAdvances(OfficeTextDirection direction) {
        const string source = "office";
        byte[] data = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "SourceSansPro-Regular.otf"));
        IOfficeFontProgram font = Assert.Single(new OfficeFontFaceCollection().Add("CFF advance", data).Faces).Program;
        var request = new OfficeTextShapingRequest(source, "CFF advance", data, true, font.UnitsPerEm,
            direction, "en", OfficeTextFeatureSettings.Default.With("kern", 1).With("liga", 1));

        OfficeTextShapingResult result = Assert.IsType<OfficeTextShapingResult>(OfficeManagedTextShapingProvider.Instance.ShapeText(request));

        // Retained same-font reference shaping has c/e=-21. Its final e origin
        // is 1800 rather than the unkerned 1821; paragraph RTL leaves this Latin
        // directional run in its natural order. The source includes ff once.
        var reference = new OfficeTextShapingResult(new[] {
            new OfficeShapedGlyph(42, "o", 0, 542),
            new OfficeShapedGlyph(514, "ff", 1, 577),
            new OfficeShapedGlyph(36, "i", 3, 246),
            new OfficeShapedGlyph(30, "c", 4, 435),
            new OfficeShapedGlyph(32, "e", 5, 496)
        }, direction);
        Assert.Equal(reference.Glyphs.Select(glyph => glyph.GlyphId), result.Glyphs.Select(glyph => glyph.GlyphId));
        Assert.Equal(source, string.Concat(result.Glyphs.OrderBy(glyph => glyph.TextIndex).Select(glyph => glyph.UnicodeText)));
        Assert.Equal(2296D, font.MeasureShapedText(source, result, font.UnitsPerEm));
        Assert.Equal(2296D, font.MeasureShapedText(source, reference, font.UnitsPerEm));

        List<List<OfficePoint>> expected = font.GetShapedTextContours(source, reference, 0D, 0D, font.UnitsPerEm);
        List<List<OfficePoint>> actual = font.GetShapedTextContours(source, result, 0D, 0D, font.UnitsPerEm);
        Assert.NotEmpty(actual);
        Assert.Equal(expected.Count, actual.Count);
        for (int index = 0; index < expected.Count; index++) Assert.Equal(expected[index].ToArray(), actual[index].ToArray());
    }
}
