using System.Text.Json;

namespace OfficeIMO.Bibliography.Tests;

public sealed class CslNumericContractTests {
    [Theory]
    [InlineData("volume", "ordinal", "2", "2n&7-8")]
    [InlineData("edition", "ordinal", "2 & 3", "2n&7-8 & 3n&7-8")]
    [InlineData("volume", "long-ordinal", "2", "word&7-8")]
    [InlineData("locator", "ordinal", "2-3", "2n&7-8–3n&7-8")]
    [InlineData("locator", "long-ordinal", "2-3", "word&7-8–third")]
    public void GeneratedOrdinalTermsRemainLiteral(string variable, string form, string value, string expected) {
        const string locale = "<locale><terms><term name=\"ordinal\">n&amp;7-8</term><term name=\"long-ordinal-02\">word&amp;7-8</term></terms></locale>";
        foreach (CslOutputFormat output in new[] { CslOutputFormat.PlainText, CslOutputFormat.Html })
            Assert.Equal(expected, System.Net.WebUtility.HtmlDecode(Render(variable, value,
                "<number variable=\"" + variable + "\" form=\"" + form + "\"/>", locale, output)));
    }

    [Theory]
    [InlineData("page", "i-ix", "pages")]
    [InlineData("locator", "I–IV", "pages")]
    [InlineData("page", "ix & 2", "pages")]
    [InlineData("volume", "ii, iv", "volumes")]
    [InlineData("page", "i", "page")]
    [InlineData("page", "Michaelson-Morley", "page")]
    [InlineData("page", "3-B", "page")]
    [InlineData("volume", "2", "volume")]
    [InlineData("number-of-volumes", "2", "volumes")]
    public void ContextualLabelsRecognizeNumeralListsWithoutTreatingWordsAsRanges(string variable, string value, string expected) =>
        Assert.Equal(expected, Render(variable, value, "<label variable=\"" + variable + "\"/>"));

    [Theory]
    [InlineData("text")]
    [InlineData("number")]
    public void LiteralLocatorHyphensUseTheRangeDelimiterForBothRenderingElements(string element) =>
        Assert.Equal("chapter I–V", Render("locator", "chapter I-V", "<" + element + " variable=\"locator\"/>"));

    [Theory]
    [InlineData("always", "i", "pages")]
    [InlineData("never", "i-ix", "page")]
    public void ExplicitPluralityOverridesNumeralDetection(string plural, string value, string expected) =>
        Assert.Equal(expected, Render("page", value, "<label variable=\"page\" plural=\"" + plural + "\"/>"));

    [Fact]
    public void GenderAndOrdinalNamespacesRemainIndependent() {
        const string locale = "<locale><terms><term name=\"edition\" gender=\"feminine\">edition</term><term name=\"issue\" gender=\"masculine\">issue</term>" +
            "<term name=\"ordinal\">n</term><term name=\"ordinal-01\" gender-form=\"feminine\">f</term>" +
            "<term name=\"ordinal-01\" gender-form=\"masculine\">m</term><term name=\"ordinal-02\" gender-form=\"masculine\">mm</term></terms></locale>";
        Assert.Equal("1f, 2n, 3n", Render("edition", "1,2,3", "<number variable=\"edition\" form=\"ordinal\"/>", locale));
        Assert.Equal("1m, 2mm, 3n", Render("issue", "1,2,3", "<number variable=\"issue\" form=\"ordinal\"/>", locale));
    }

    [Theory]
    [InlineData("ordinal-11", "", "11", "11x")]
    [InlineData("ordinal-11", "", "111", "111x")]
    [InlineData("ordinal-01", "", "11", "11x")]
    [InlineData("ordinal-01", " match=\"whole-number\"", "11", "11")]
    public void PartialModernOrdinalSetsKeepTheirMatchingRules(string term, string attributes, string value, string expected) =>
        Assert.Equal(expected, Render("volume", value, "<number variable=\"volume\" form=\"ordinal\"/>",
            "<locale><terms><term name=\"" + term + "\"" + attributes + ">x</term></terms></locale>"));

    [Fact]
    public void TheCompleteLegacySuffixSetRetainsTeenAndLastDigitRules() =>
        Assert.Equal("11d, 21a, 22b, 23c, 24d", Render("volume", "11,21,22,23,24", "<number variable=\"volume\" form=\"ordinal\"/>",
            "<locale><terms><term name=\"ordinal-01\">a</term><term name=\"ordinal-02\">b</term>" +
            "<term name=\"ordinal-03\">c</term><term name=\"ordinal-04\">d</term></terms></locale>"));

    [Fact]
    public void AnExplicitEmptyLongOrdinalDoesNotFallBack() =>
        Assert.Equal(string.Empty, Render("volume", "1", "<number variable=\"volume\" form=\"long-ordinal\"/>",
            "<locale><terms><term name=\"long-ordinal-01\"/></terms></locale>"));

    [Theory]
    [InlineData("locator", "ordinal", "2-й")]
    [InlineData("locator", "long-ordinal", "второй")]
    [InlineData("number-of-volumes", "ordinal", "2-й")]
    [InlineData("number-of-volumes", "long-ordinal", "второй")]
    public void LocatorAndCountOrdinalsUseTheAccompanyingNounGender(string variable, string form, string expected) =>
        Assert.Equal(expected, Render(variable, "2", "<number variable=\"" + variable + "\" form=\"" + form + "\"/>",
            locatorType: "volume", defaultLocale: "ru-RU"));

    [Theory]
    [InlineData("locator", "ordinal", "2-я")]
    [InlineData("locator", "long-ordinal", "вторая")]
    [InlineData("number-of-pages", "ordinal", "2-я")]
    [InlineData("number-of-pages", "long-ordinal", "вторая")]
    public void PageNounOverridesApplyToBothLocatorsAndCounts(string variable, string form, string expected) =>
        Assert.Equal(expected, Render(variable, "2", "<number variable=\"" + variable + "\" form=\"" + form + "\"/>",
            "<locale><terms><term name=\"page\" gender=\"feminine\">page</term></terms></locale>", defaultLocale: "ru-RU"));

    private static string Render(string variable, string value, string layout, string locale = "", CslOutputFormat output = CslOutputFormat.PlainText,
        string locatorType = "page", string? defaultLocale = null) {
        string fields = variable == "locator" ? string.Empty : "," + JsonSerializer.Serialize(variable) + ":" + JsonSerializer.Serialize(value);
        BibliographyDocument document = BibliographyDocument.Parse("[{\"id\":\"one\",\"type\":\"book\"" + fields + "}]", BibliographyFormat.CslJson).Document;
        CslStyle style = CslStyle.Parse("<style xmlns=\"http://purl.org/net/xbiblio/csl\" version=\"1.0\" class=\"in-text\"" +
            (defaultLocale == null ? string.Empty : " default-locale=\"" + defaultLocale + "\"") + ">" + locale +
            "<citation><layout>" + layout + "</layout></citation></style>");
        var citation = new CslCitation("cite");
        citation.Items.Add(new CslCitationItem("one") { Locator = variable == "locator" ? value : null, LocatorType = locatorType });
        return new CslProcessor(document, style, new CslRenderOptions { OutputFormat = output }).Render(new[] { citation }).Citations.Single().Content;
    }
}
