using System.Text.Json;

namespace OfficeIMO.Bibliography.Tests;

public sealed class CslNumericScanContractTests {
    [Theory]
    [InlineData("A١B", "yes")]
    [InlineData("Ⅳ2", "no")]
    [InlineData("2²", "no")]
    [InlineData("2\u0301", "no")]
    [InlineData("\u2003É12β\u0085,\t34\u00a0", "yes")]
    [InlineData("１２３", "yes")]
    [InlineData("1 2", "no")]
    [InlineData("2, ", "no")]
    public void NumericConditionsPreserveUnicodeCategoriesAndCompleteSegments(string value, string expected) =>
        Assert.Equal(expected, Render("volume", value, NumericCondition));

    [Theory]
    [InlineData("٣, 12", "٣, 12th")]
    [InlineData("é2β, 12", "é2β, 12th")]
    [InlineData("１２, 12", "１２, 12th")]
    public void OrdinalsConvertBareMachineNumeralsWithoutChangingSourceAffixesOrDigitGlyphs(string value, string expected) =>
        Assert.Equal(expected, Render("volume", value, "<number variable=\"volume\" form=\"ordinal\"/>"));

    [Theory]
    [InlineData("volume", "1st, 2nd & 3rd-4th")]
    [InlineData("locator", "1st, 2nd & 3rd–4th")]
    public void UnicodeSeparatorWhitespaceUsesTheExistingNumberAndLocatorPolicies(string variable, string expected) =>
        Assert.Equal(expected, Render(variable, "1\t,\u00852\u2003&\u202f3\r\n-\u00a04", "<number variable=\"" + variable + "\" form=\"ordinal\"/>"));

    [Theory]
    [InlineData("A1-2B", "A1–2B")]
    [InlineData("٣-٤", "٣–٤")]
    [InlineData("Ⅳ-Ⅴ", "Ⅳ-Ⅴ")]
    [InlineData("1 - 2", "1 - 2")]
    public void TextNumberHyphensRequireAdjacentDecimalDigits(string value, string expected) =>
        Assert.Equal(expected, Render("volume", value, "<text variable=\"volume\"/>"));

    [Fact]
    public void LocatorConnectorLocalizationRetainsEntitiesAndDoesNotRescanTheLocaleTerm() {
        const string locale = "<locale><terms><term name=\"and\" form=\"symbol\">PLUS7-8</term></terms></locale>";
        Assert.Equal("A & B PLUS7-8 C & D & E &unknown; F PLUS7-8a; G",
            Render("locator", "A &amp; B & C &#38; D &#x26; E &unknown; F &a; G", "<text variable=\"locator\"/>", locale: locale));
    }

    [Theory]
    [InlineData("IIV-III", "page")]
    [InlineData("IIII-IV", "page")]
    [InlineData("MMMCMXCIX-iii", "pages")]
    public void RomanLabelsRequireCanonicalNumerals(string value, string expected) =>
        Assert.Equal(expected, Render("page", value, "<label variable=\"page\"/>"));

    [Theory]
    [InlineData("123-128-132", "123–8-132")]
    [InlineData("_123-128_", "_123–8_")]
    [InlineData("A - - B123-B128", "A - - B123–8")]
    [InlineData("i-ix, IIV-III, MMMCMXCIX-iii", "i–ix, IIV-III, MMMCMXCIX–iii")]
    public void PageRangeScanningRetainsNonoverlappingRangesAndMalformedSource(string value, string expected) =>
        Assert.Equal(expected, Render("page", value, "<text variable=\"page\"/>", styleOptions: "page-range-format=\"minimal\""));

    [Theory]
    [InlineData("volume", "1,2", "numeric", 3)]
    [InlineData("volume", "1,2", "ordinal", 6)]
    [InlineData("page", "23-8", "roman", 10)]
    public void NumberExpansionEnforcesTheIntermediateLimit(string variable, string value, string form, int maximum) =>
        Assert.Contains("MaximumIntermediateCharacters", Assert.Throws<InvalidDataException>(() =>
            Render(variable, value, "<number variable=\"" + variable + "\" form=\"" + form + "\"/>",
                styleOptions: "page-range-format=\"expanded\"", maximum: maximum)).Message);

    [Theory]
    [InlineData("volume", "<number variable=\"volume\"/>")]
    [InlineData("locator", "<text variable=\"locator\"/>")]
    [InlineData("page", "<text variable=\"page\"/>")]
    public void ALongLocaleConnectorCannotBypassTheIntermediateLimit(string variable, string layout) {
        string locale = "<locale><terms><term name=\"and\" form=\"symbol\">" + new string('x', 64) + "</term></terms></locale>";
        Assert.Contains("MaximumIntermediateCharacters", Assert.Throws<InvalidDataException>(() =>
            Render(variable, "1&2", layout, locale: locale, maximum: 10)).Message);
    }

    [Theory]
    [InlineData(false, "yes")]
    [InlineData(true, "no")]
    public void LongNumericMetadataCanBeClassifiedWithoutRenderingTheWholeSource(bool invalidTail, string expected) =>
        Assert.Equal(expected, Render("volume", new string('9', 20000) + (invalidTail ? "x y" : string.Empty), NumericCondition, maximum: 4));

    [Fact]
    public void LongWhitespaceRunsAreNormalizedWithinTheRenderedOutputLimit() =>
        Assert.Equal("1, 2", Render("volume", "1" + new string(' ', 20000) + "," + new string('\u2003', 20000) + "2",
            "<number variable=\"volume\"/>", maximum: 4));

    private const string NumericCondition = "<choose><if is-numeric=\"volume\"><text value=\"yes\"/></if><else><text value=\"no\"/></else></choose>";

    private static string Render(string variable, string value, string layout, string styleOptions = "", string locale = "", int maximum = 65536) {
        string fields = variable == "locator" ? string.Empty : "," + JsonSerializer.Serialize(variable) + ":" + JsonSerializer.Serialize(value);
        BibliographyDocument data = BibliographyDocument.Parse("[{\"id\":\"one\",\"type\":\"book\"" + fields + "}]", BibliographyFormat.CslJson).Document;
        CslStyle style = CslStyle.Parse("<style xmlns=\"http://purl.org/net/xbiblio/csl\" version=\"1.0\" class=\"in-text\" " +
            styleOptions + ">" + locale + "<citation><layout>" + layout + "</layout></citation></style>");
        var citation = new CslCitation("cite");
        citation.Items.Add(new CslCitationItem("one") { Locator = variable == "locator" ? value : null });
        return new CslProcessor(data, style, new CslRenderOptions { MaximumIntermediateCharacters = maximum }).Render(new[] { citation }).Citations.Single().Content;
    }
}
