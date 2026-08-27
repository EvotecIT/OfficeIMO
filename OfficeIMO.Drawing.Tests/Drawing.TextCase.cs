using System.Globalization;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed class OfficeTextCaseTests {
    [Theory]
    [InlineData(OfficeTextCase.None, "mIXed API text", "mIXed API text")]
    [InlineData(OfficeTextCase.Uppercase, "Mixed api text", "MIXED API TEXT")]
    [InlineData(OfficeTextCase.Lowercase, "Mixed API Text", "mixed api text")]
    [InlineData(OfficeTextCase.TitleCase, "mIXed api TEXT", "Mixed Api Text")]
    [InlineData(OfficeTextCase.SentenceCase, "hELLO WORLD. aNOTHER SENTENCE! final QUESTION?", "Hello world. Another sentence! Final question?")]
    [InlineData(OfficeTextCase.ToggleCase, "Mixed API 42", "mIXED api 42")]
    public void ApplyTransformsTextWithoutFormatDependencies(OfficeTextCase textCase, string input, string expected) {
        Assert.Equal(expected, OfficeTextCaseTransformer.Apply(input, textCase, CultureInfo.InvariantCulture));
    }

    [Fact]
    public void ApplyUsesSelectedCulture() {
        CultureInfo culture = CultureInfo.GetCultureInfo("tr-TR");
        Assert.Equal("İSTANBUL", OfficeTextCaseTransformer.Apply("istanbul", OfficeTextCase.Uppercase, culture));
        Assert.Equal("ıstanbul", OfficeTextCaseTransformer.Apply("ISTANBUL", OfficeTextCase.Lowercase, culture));
    }

    [Fact]
    public void RichTextRunCopyPreservesDrawingStyle() {
        OfficeRichTextRun source = new("Styled", 14, OfficeColor.FromRgb(51, 102, 153),
            bold: true, italic: true, fontFamily: "Aptos",
            backgroundColor: OfficeColor.FromRgb(240, 240, 240),
            underlineStyle: OfficeTextDecorationStyle.Dashed,
            strikethroughStyle: OfficeTextDecorationStyle.Double,
            baseline: OfficeTextBaseline.Superscript);

        OfficeRichTextRun actual = source.WithTextCase(OfficeTextCase.ToggleCase);

        Assert.Equal("sTYLED", actual.Text);
        Assert.Equal(source.FontSize, actual.FontSize);
        Assert.Equal(source.Color, actual.Color);
        Assert.Equal(source.Bold, actual.Bold);
        Assert.Equal(source.Italic, actual.Italic);
        Assert.Equal(source.Underline, actual.Underline);
        Assert.Equal(source.Strikethrough, actual.Strikethrough);
        Assert.Equal(source.FontFamily, actual.FontFamily);
        Assert.Equal(source.BackgroundColor, actual.BackgroundColor);
        Assert.Equal(source.UnderlineStyle, actual.UnderlineStyle);
        Assert.Equal(source.StrikethroughStyle, actual.StrikethroughStyle);
        Assert.Equal(source.Baseline, actual.Baseline);
    }
}
