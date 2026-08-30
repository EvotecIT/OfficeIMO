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

    [Fact]
    public void FontFallbackSelectsTheFaceThatSupportsTheRequestedVariationSequence() {
        const string sequence = "\u2764\uFE0F";
        Assert.True(OfficeFontUnicodeRangeSet.TryParseCss("U+2764", out OfficeFontUnicodeRangeSet? narrow));
        Assert.True(OfficeFontUnicodeRangeSet.TryParseCss("U+2700-27BF", out OfficeFontUnicodeRangeSet? broad));
        var fonts = new OfficeFontFaceCollection()
            .Add("Scoped", ManagedTextShapingTestAssets.CreateFont(0x2764), OfficeFontStyle.Regular, narrow!)
            .Add(
                "Scoped",
                ManagedTextShapingTestAssets.CreateFontWithVariationSequence(0x2764, 0xFE0F),
                OfficeFontStyle.Regular,
                broad!);

        OfficeFontFallbackRun run = Assert.Single(fonts.PlanFallbackRuns(sequence, "Scoped"));

        Assert.Equal(sequence, run.Text);
        Assert.Equal(fonts.Faces[1].ResourceFamilyName, run.FamilyName);
        Assert.NotEqual(fonts.Faces[0].ResourceFamilyName, run.FamilyName);
    }

    [Fact]
    public void ManagedFontMapsANonDefaultVariationGlyphForRendering() {
        const string sequence = "\u2764\uFE0F";
        OfficeTrueTypeFont font = Assert.IsType<OfficeTrueTypeFont>(OfficeTrueTypeFont.TryLoad(
            ManagedTextShapingTestAssets.CreateFontWithNonDefaultVariationSequence(0x2764, 0xFE0F)));

        Assert.True(font.HasGlyphs(sequence));
        double baseMaximumX = font.GetTextContours("\u2764", 0, 0, 1000).SelectMany(contour => contour).Max(point => point.X);
        double variationMaximumX = font.GetTextContours(sequence, 0, 0, 1000).SelectMany(contour => contour).Max(point => point.X);
        Assert.True(variationMaximumX > baseMaximumX);
    }

    [Fact]
    public void VariationCoverageRejectsFormat14OutsideTheUnicodeUvsEncodingRecord() {
        const string sequence = "\u2764\uFE0F";
        OfficeTrueTypeFont font = Assert.IsType<OfficeTrueTypeFont>(OfficeTrueTypeFont.TryLoad(
            ManagedTextShapingTestAssets.CreateFontWithMistypedVariationSequenceRecord(0x2764, 0xFE0F)));

        Assert.False(font.HasGlyphs(sequence));
    }

    [Fact]
    public void VariationCoverageSupportsLargeValidNonDefaultMappingTables() {
        const string sequence = "\u2764\uFE0F";
        OfficeTrueTypeFont font = Assert.IsType<OfficeTrueTypeFont>(OfficeTrueTypeFont.TryLoad(
            ManagedTextShapingTestAssets.CreateFontWithLargeNonDefaultVariationSequence(0x2764, 0xFE0F)));

        Assert.True(font.HasGlyphs(sequence));
        double baseMaximumX = font.GetTextContours("\u2764", 0, 0, 1000).SelectMany(contour => contour).Max(point => point.X);
        double variationMaximumX = font.GetTextContours(sequence, 0, 0, 1000).SelectMany(contour => contour).Max(point => point.X);
        Assert.True(variationMaximumX > baseMaximumX);
    }
}
