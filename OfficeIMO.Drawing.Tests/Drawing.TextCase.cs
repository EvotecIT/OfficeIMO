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
    [InlineData(OfficeTextCase.Capitalize, "iPhone eBOOK don't-stop", "IPhone EBOOK Don't-Stop")]
    public void ApplyTransformsTextWithoutFormatDependencies(OfficeTextCase textCase, string input, string expected) {
        Assert.Equal(expected, OfficeTextCaseTransformer.Apply(input, textCase, CultureInfo.InvariantCulture));
    }

    [Fact]
    public void ApplySegmentsPreservesSentenceAndWordContextAcrossFormattingBoundaries() {
        IReadOnlyList<string> sentence = OfficeTextCaseTransformer.ApplySegments(
            new[] { "hELLO. ", "aNOTHER", " SENTENCE" },
            OfficeTextCase.SentenceCase,
            CultureInfo.InvariantCulture);
        IReadOnlyList<string> title = OfficeTextCaseTransformer.ApplySegments(
            new[] { "mIXed ", "caSE" },
            OfficeTextCase.TitleCase,
            CultureInfo.InvariantCulture);

        Assert.Equal(new[] { "Hello. ", "Another", " sentence" }, sentence);
        Assert.Equal(new[] { "Mixed ", "Case" }, title);
    }

    [Fact]
    public void ApplySegmentsTransformsManyBoundariesInOneContextPreservingPass() {
        string[] segments = Enumerable.Range(0, 10000)
            .Select(index => index == 5000 ? "mixed\uE000case" : "a")
            .ToArray();

        IReadOnlyList<string> transformed = OfficeTextCaseTransformer.ApplySegments(
            segments,
            OfficeTextCase.Capitalize,
            CultureInfo.InvariantCulture);

        Assert.Equal(segments.Length, transformed.Count);
        Assert.Equal(
            OfficeTextCaseTransformer.Apply(string.Concat(segments), OfficeTextCase.Capitalize, CultureInfo.InvariantCulture),
            string.Concat(transformed));
        Assert.Equal("mixed\uE000case", transformed[5000]);
    }

    [Theory]
    [InlineData("123abc")]
    [InlineData("foo+bar_baz/qux")]
    [InlineData("don't-stop")]
    [InlineData("\u01F3uro")]
    public void TitleCaseMatchesTextInfoWordAndTitleLetterSemantics(string input) {
        CultureInfo culture = CultureInfo.InvariantCulture;
        Assert.Equal(
            culture.TextInfo.ToTitleCase(input.ToLower(culture)),
            OfficeTextCaseTransformer.Apply(input, OfficeTextCase.TitleCase, culture));
    }

    [Fact]
    public void ApplyUsesSelectedCulture() {
        CultureInfo culture = CultureInfo.GetCultureInfo("tr-TR");
        Assert.Equal("İSTANBUL", OfficeTextCaseTransformer.Apply("istanbul", OfficeTextCase.Uppercase, culture));
        Assert.Equal("ıstanbul", OfficeTextCaseTransformer.Apply("ISTANBUL", OfficeTextCase.Lowercase, culture));
    }

    [Fact]
    public void ApplySegmentsKeepsCultureSpecificTitleCaseAcrossRunBoundaries() {
        CultureInfo culture = CultureInfo.GetCultureInfo("nl-NL");
        IReadOnlyList<string> transformed = OfficeTextCaseTransformer.ApplySegments(
            new[] { "i", "jsselmeer" },
            OfficeTextCase.TitleCase,
            culture);
        string expected = culture.TextInfo.ToTitleCase("ijsselmeer");

        Assert.Equal("I", transformed[0]);
        Assert.Equal(expected.Substring(1), transformed[1]);
        Assert.Equal(expected, string.Concat(transformed));
    }

    [Fact]
    public void ApplySegmentsKeepsCrossBoundaryGraphemeCodePointsInTheirSourceRuns() {
        IReadOnlyList<string> combining = OfficeTextCaseTransformer.ApplySegments(
            new[] { "e", "\u0301" },
            OfficeTextCase.Uppercase,
            CultureInfo.InvariantCulture);
        IReadOnlyList<string> emoji = OfficeTextCaseTransformer.ApplySegments(
            new[] { "\U0001F469", "\u200D", "\U0001F4BB" },
            OfficeTextCase.ToggleCase,
            CultureInfo.InvariantCulture);

        Assert.Equal(new[] { "E", "\u0301" }, combining);
        Assert.Equal(new[] { "\U0001F469", "\u200D", "\U0001F4BB" }, emoji);
    }

    [Fact]
    public void ApplyTransformsSupplementaryUnicodeLettersAsWholeTextElements() {
        const string deseretCapitalLongI = "\U00010400";
        const string deseretSmallLongI = "\U00010428";

        Assert.Equal(deseretSmallLongI,
            OfficeTextCaseTransformer.Apply(deseretCapitalLongI, OfficeTextCase.ToggleCase, CultureInfo.InvariantCulture));
        Assert.Equal(deseretCapitalLongI + "bc. " + deseretCapitalLongI + "ef",
            OfficeTextCaseTransformer.Apply(
                deseretSmallLongI + "BC. " + deseretSmallLongI + "EF",
                OfficeTextCase.SentenceCase,
                CultureInfo.InvariantCulture));
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
