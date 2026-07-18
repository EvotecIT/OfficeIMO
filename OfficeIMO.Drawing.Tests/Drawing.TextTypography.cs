using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public class DrawingTextTypographyTests {
    [Fact]
    public void TextElements_ResolveBaseDirectionFromFirstStrongCharacter() {
        Assert.Equal(OfficeTextDirection.RightToLeft, OfficeTextElements.ResolveBaseDirection("123 سلام"));
        Assert.Equal(OfficeTextDirection.LeftToRight, OfficeTextElements.ResolveBaseDirection("123 Office"));
        Assert.Equal(OfficeTextDirection.Auto, OfficeTextElements.ResolveBaseDirection("123 -"));
        Assert.Equal(OfficeTextDirection.LeftToRight, OfficeTextElements.ResolveBaseDirection("\u0661\u0662 Office"));
        Assert.Equal(OfficeTextDirection.RightToLeft, OfficeTextElements.ResolveBaseDirection("\u200F123"));
        Assert.Equal(OfficeTextDirection.RightToLeft, OfficeTextElements.ResolveBaseDirection("\u0800"));
        Assert.Equal(OfficeTextDirection.RightToLeft, OfficeTextElements.ResolveBaseDirection("\u0870"));
        Assert.Equal(OfficeTextDirection.RightToLeft, OfficeTextElements.ResolveBaseDirection("\U00010900"));
    }

    [Fact]
    public void LineBreaks_KeepCjkPunctuationAndGraphemeClustersTogether() {
        const string cjk = "日本東京、大阪京都";
        IReadOnlyList<int> cjkBreaks = OfficeTextLineBreaks.GetBreakPositions(cjk);

        int punctuationIndex = cjk.IndexOf('、');
        Assert.DoesNotContain(punctuationIndex, cjkBreaks);
        Assert.Contains(punctuationIndex + 1, cjkBreaks);

        const string combining = "漢e\u0301字";
        int combiningMarkIndex = combining.IndexOf('\u0301');
        Assert.False(OfficeTextLineBreaks.IsValidBreakPosition(combining, combiningMarkIndex));
        Assert.DoesNotContain(combiningMarkIndex, OfficeTextLineBreaks.GetBreakPositions(combining));

        const string supplementary = "漢\U00020000字";
        Assert.False(OfficeTextLineBreaks.IsValidBreakPosition(supplementary, 2));
        Assert.DoesNotContain(2, OfficeTextLineBreaks.GetBreakPositions(supplementary));
    }

    [Fact]
    public void LineBreaks_ExposeUsefulNonCjkTokenBoundaries() {
        Assert.Equal(new[] { 6, 11 }, OfficeTextLineBreaks.GetBreakPositions("alpha-beta/gamma"));
        Assert.True(OfficeTextLineBreaks.IsValidBreakPosition("alpha-beta", 6));
        Assert.False(OfficeTextLineBreaks.IsValidBreakPosition("alpha-beta", 0));
    }

    [Fact]
    public void TextLayout_UsesSharedCjkBreakRulesInsteadOfStartingWithClosingPunctuation() {
        const string text = "日本東京、大阪京都";

        IReadOnlyList<OfficeTextLine> lines = OfficeTextLayoutEngine.WrapLines(
            text,
            fontSize: 1,
            maxWidth: 4,
            (value, _) => value?.Length ?? 0);

        Assert.True(lines.Count > 1);
        Assert.Equal(text, string.Concat(lines.Select(line => line.Text)));
        Assert.DoesNotContain(lines.Skip(1), line => line.Text.StartsWith("、", StringComparison.Ordinal));
    }

    [Fact]
    public void TextLigatures_PreferLongestStandardLatinMatch() {
        Assert.True(OfficeTextLigatures.TryGetLatinPresentationForm("office", 1, out int scalar, out int length));
        Assert.Equal(0xFB03, scalar);
        Assert.Equal(3, length);
        Assert.False(OfficeTextLigatures.TryGetLatinPresentationForm("office", 5, out scalar, out length));
        Assert.Equal(0, scalar);
        Assert.Equal(0, length);
    }

    [Fact]
    public void TextShapingContracts_AreImmutableAndValidateMappings() {
        byte[] fontData = { 1, 2, 3 };
        var request = new OfficeTextShapingRequest(
            "AB",
            "Example",
            fontData,
            isOpenTypeCff: false,
            unitsPerEm: 2048,
            OfficeTextDirection.LeftToRight,
            "en-US");
        fontData[0] = 9;
        byte[] firstSnapshot = request.FontData;
        firstSnapshot[1] = 9;

        Assert.Equal(new byte[] { 1, 2, 3 }, request.FontData);
        Assert.Equal(2048, request.UnitsPerEm);
        Assert.Equal("en-US", request.Language);

        var result = new OfficeTextShapingResult(new[] {
            new OfficeShapedGlyph(12, "A", 0, 1024),
            new OfficeShapedGlyph(13, "B", 1, 1024, 10, -5)
        });
        Assert.Equal(2, result.Glyphs.Count);
        Assert.IsNotType<OfficeShapedGlyph[]>(result.Glyphs);
        Assert.Equal(10, result.Glyphs[1].OffsetX);
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeShapedGlyph(0, "A", 0));
        Assert.Throws<ArgumentException>(() => new OfficeShapedGlyph(1, string.Empty, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeTextShapingRequest("A", "Example", new byte[] { 1 }, false, 0));
    }
}
