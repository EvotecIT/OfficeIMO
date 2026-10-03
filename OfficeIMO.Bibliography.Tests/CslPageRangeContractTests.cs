using System.Text.Json;

namespace OfficeIMO.Bibliography.Tests;

public sealed class CslPageRangeContractTests {
    [Theory]
    [InlineData("expanded")]
    [InlineData("minimal")]
    [InlineData("minimal-two")]
    [InlineData("chicago")]
    [InlineData("chicago-15")]
    [InlineData("chicago-16")]
    public void LongerEndpointsRetainAllDigits(string format) =>
        Assert.Equal("12991–123001, 999–1001, 92–113", Render("12991-123001, 999-1001, 92-113", format));

    [Theory]
    [InlineData("expanded", "n11564–n11568, 8n11564–8n11568, S213–S235")]
    [InlineData("minimal", "n11564–8, 8n11564–8, S213–35")]
    [InlineData("minimal-two", "n11564–68, 8n11564–68, S213–35")]
    [InlineData("chicago", "n11564–68, 8n11564–68, S213–35")]
    [InlineData("chicago-15", "n11564–68, 8n11564–68, S213–35")]
    [InlineData("chicago-16", "n11564–68, 8n11564–68, S213–35")]
    public void MatchingPrefixesPermitRangeFormatting(string format, string expected) =>
        Assert.Equal(expected, Render("n11564 - n1568, 8n11564-8n1568, S213-S235", format));

    [Fact]
    public void DifferentPrefixesRetainIdentifierMeaningWhileRangeWhitespaceIsNormalized() =>
        Assert.Equal("N110-5, 110-N6, N110-P5, 123N110-N5, 456K200-99, 000c23-22",
            Render("N110 - 5, 110 - N6, N110 - P5, 123N110 - N5, 456K200 - 99, 000c23 - 22", "expanded"));

    [Fact]
    public void RomanEndpointsUseThePageDelimiterWithoutAbbreviation() =>
        Assert.Equal("xxv–xxviii, cvi–cix, I–IV; article-title",
            Render("xxv - xxviii, cvi-cix, I-IV; article-title", "minimal"));

    [Fact]
    public void RomanNumberRenderingStillUsesThePageDelimiter() =>
        Assert.Equal("xxiii–xxviii", Render("23-28", "minimal", "<number variable=\"page\" form=\"roman\"/>"));

    [Fact]
    public void NumericSuffixesRemainAttachedToEachEndpoint() =>
        Assert.Equal("S213a–S235a, 23a-23c", Render("S213a - S235a, 23a - 23c", "minimal"));

    [Fact]
    public void NumericNumberRenderingAndTextShareTheRangeRules() =>
        Assert.Equal("n11564–68, 12991–123001", Render("n11564 - n1568, 12991-123001", "chicago", "<number variable=\"page\"/>"));

    [Fact]
    public void ALocaleCanOverrideTheRangeDelimiter() =>
        Assert.Equal("S213 to 35, xxv to xxviii", Render("S213-S235, xxv-xxviii", "minimal",
            locale: "<locale><terms><term name=\"page-range-delimiter\"> to </term></terms></locale>"));

    [Theory]
    [InlineData(null, "N110 - N5, xxv-xxviii, 110 - 5")]
    [InlineData("expanded", "N110–N115, xxv–xxviii, 110–115")]
    public void TextPageFormattingIsOptIn(string? format, string expected) =>
        Assert.Equal(expected, Render("N110 - N5, xxv-xxviii, 110 - 5", format));

    [Theory]
    [InlineData("chicago", "3–10, 100–104, 107–8, 321–25, 1496–1504, 2787–2816")]
    [InlineData("chicago-15", "3–10, 100–104, 107–8, 321–25, 1496–1504, 2787–2816")]
    [InlineData("chicago-16", "3–10, 100–104, 107–8, 321–25, 1496–504, 2787–816")]
    public void ChicagoEditionRulesRemainDistinct(string format, string expected) =>
        Assert.Equal(expected, Render("3-10, 100-104, 107-108, 321-325, 1496-1504, 2787-2816", format));

    [Theory]
    [InlineData("expanded", "roman", "123-8", "cxxiii–cxxviii")]
    [InlineData("minimal", "roman", "123-8", "cxxiii–cxxviii")]
    [InlineData("expanded", "ordinal", "123-8", "123rd–128th")]
    [InlineData("expanded", "long-ordinal", "123-8", "123rd–128th")]
    [InlineData("expanded", "ordinal", "23-28", "23rd–28th")]
    [InlineData("expanded", "roman", "3998-4001", "mmmcmxcviii–4001")]
    public void PageEndpointsAreResolvedBeforeTheirNumeralFormIsApplied(string format, string form, string page, string expected) {
        foreach (CslOutputFormat output in new[] { CslOutputFormat.PlainText, CslOutputFormat.Html })
            Assert.Equal(expected, System.Net.WebUtility.HtmlDecode(Render(page, format, "<number variable=\"page\" form=\"" + form + "\"/>", output: output)));
    }

    [Fact]
    public void GeneratedOrdinalSuffixesDoNotHideTheLocalizedPageDelimiter() =>
        Assert.Equal("123rd to 128th", Render("123-8", "expanded", "<number variable=\"page\" form=\"ordinal\"/>",
            "<locale><terms><term name=\"page-range-delimiter\"> to </term></terms></locale>"));

    [Theory]
    [InlineData("１２３-１２８", "minimal", "１２３–８")]
    [InlineData("１２３-８", "expanded", "１２３–１２８")]
    [InlineData("۱۲۳-۱۲۸", "minimal-two", "۱۲۳–۲۸")]
    [InlineData("۱۲۳-۸", "expanded", "۱۲۳–۱۲۸")]
    [InlineData("１００-１０４, １０７-１０８", "chicago-16", "１００–１０４, １０７–８")]
    public void DecimalDigitGlyphsSurviveRangeFormatting(string page, string format, string expected) {
        Assert.Equal(expected, Render(page, format));
        Assert.Equal(expected, Render(page, format, "<number variable=\"page\"/>"));
    }

    [Theory]
    [InlineData("roman", "to", "cxxiiitocxxviii")]
    [InlineData("ordinal", "to", "123rdto128th")]
    [InlineData("long-ordinal", "to", "123rdto128th")]
    [InlineData("ordinal", " – ", "123rd – 128th")]
    [InlineData("ordinal", " - ", "123rd - 128th")]
    public void LocaleDelimiterTextIsNotReinterpretedAsSourceNumberAffixes(string form, string delimiter, string expected) {
        string locale = "<locale><terms><term name=\"page-range-delimiter\">" + delimiter + "</term></terms></locale>";
        foreach (CslOutputFormat output in new[] { CslOutputFormat.PlainText, CslOutputFormat.Html })
            Assert.Equal(expected, System.Net.WebUtility.HtmlDecode(Render("123-8", "expanded", "<number variable=\"page\" form=\"" + form + "\"/>", locale, output)));
    }

    [Theory]
    [InlineData("roman", "A123–8, 123a–128a, ix & x")]
    [InlineData("ordinal", "A123–8, 123a–128a, 9th & 10th")]
    public void SourceAffixesKeepTheirPagePolicyWhileOtherNumbersAreConverted(string form, string expected) =>
        Assert.Equal(expected, Render("A123-A128, 123a-128a, 9 & 10", "minimal", "<number variable=\"page\" form=\"" + form + "\"/>"));

    [Theory]
    [InlineData("<text variable=\"page\"/>", "123&8 + 150&3")]
    [InlineData("<number variable=\"page\"/>", "123&8 + 150&3")]
    [InlineData("<number variable=\"page\" form=\"ordinal\"/>", "123rd&128th + 150th&153rd")]
    [InlineData("<number variable=\"page\" form=\"roman\"/>", "cxxiii&cxxviii + cl&cliii")]
    public void LiteralRangeTermsAreDistinctFromSourceListConnectors(string element, string expected) =>
        Assert.Equal(expected, Render("123-128 & 150-153", "minimal", element,
            "<locale><terms><term name=\"page-range-delimiter\">&amp;</term><term name=\"and\" form=\"symbol\">+</term></terms></locale>"));

    private static string Render(string page, string? format, string element = "<text variable=\"page\"/>", string locale = "", CslOutputFormat output = CslOutputFormat.PlainText) {
        BibliographyDocument data = BibliographyDocument.Parse("[{\"id\":\"a\",\"type\":\"book\",\"page\":" +
            JsonSerializer.Serialize(page) + "}]", BibliographyFormat.CslJson).Document;
        CslStyle style = CslStyle.Parse("<style xmlns=\"http://purl.org/net/xbiblio/csl\" version=\"1.0\" class=\"in-text\"" +
            (format == null ? string.Empty : " page-range-format=\"" + format + "\"") + ">" + locale +
            "<citation><layout>" + element + "</layout></citation></style>");
        var citation = new CslCitation("c");
        citation.Items.Add(new CslCitationItem("a"));
        return new CslProcessor(data, style, new CslRenderOptions { OutputFormat = output }).Render(new[] { citation }).Citations.Single().Content;
    }
}
