using System.Text.Json;

namespace OfficeIMO.Bibliography.Tests;

public sealed class CslValueContractTests {
    private const string Header = "<style xmlns=\"http://purl.org/net/xbiblio/csl\" version=\"1.0\" class=\"in-text\">";

    [Theory]
    [InlineData(CslOutputFormat.PlainText, true, "“Title.”")]
    [InlineData(CslOutputFormat.Html, true, "“<i>Title</i>.”")]
    [InlineData(CslOutputFormat.PlainText, false, "“Title”.")]
    [InlineData(CslOutputFormat.Html, false, "“<i>Title</i>”.")]
    public void QuotePunctuationFollowsTheLocaleAcrossFormatting(CslOutputFormat format, bool inside, string expected) =>
        Assert.Equal(expected, Render("\"title\":\"Title\"", "<text variable=\"title\" quotes=\"true\" font-style=\"italic\" suffix=\".\"/>",
            locale: "<locale><style-options punctuation-in-quote=\"" + inside.ToString().ToLowerInvariant() + "\"/></locale>", options: new CslRenderOptions { OutputFormat = format }));

    [Theory]
    [InlineData("This is 'The One'", "“This is ‘The One.’”")]
    [InlineData("My “Amazing” Title", "“My ‘Amazing’ Title.”")]
    [InlineData("My “'Amazing' and Bogus” Title", "“My ‘“Amazing” and Bogus’ Title.”")]
    [InlineData("Plato's thought", "“Plato’s thought.”")]
    public void InputQuotesAlternateInsideStyleQuotes(string title, string expected) =>
        Assert.Equal(expected, Render("\"title\":" + JsonSerializer.Serialize(title), "<text variable=\"title\" quotes=\"true\" suffix=\".\"/>"));

    [Theory]
    [InlineData(CslOutputFormat.PlainText, "It is official (“My Title”) now")]
    [InlineData(CslOutputFormat.Html, "It is official (“<i>My Title</i>”) now")]
    public void InputQuotesCanEncloseFormattedText(CslOutputFormat format, string expected) =>
        Assert.Equal(expected, Render("\"title\":\"It is official (\\\"<i>My Title</i>\\\") now\"", "<text variable=\"title\"/>", options: new CslRenderOptions { OutputFormat = format }));

    [Fact]
    public void AbbreviatedYearsAndUnpairedDoubleQuotesRemainDistinct() {
        Assert.Equal("ETFA ’09", Render("\"title\":\"ETFA '09\"", "<text variable=\"title\"/>"));
        Assert.Equal("Nation of \"Positive Obligations \"", Render("\"title\":\"Nation of \\\"Positive Obligations \\\"\"", "<text variable=\"title\"/>"));
        Assert.Equal("An \"orphan \" and “valid” quote", Render("\"title\":\"An \\\"orphan \\\" and \\\"valid\\\" quote\"", "<text variable=\"title\"/>"));
    }

    [Theory]
    [InlineData("Article A?", "“Article A?”")]
    [InlineData("Article A!", "“Article A!”")]
    [InlineData("Article A.", "“Article A.”")]
    public void QuotePunctuationSuppressesARedundantFinalPeriod(string title, string expected) =>
        Assert.Equal(expected, Render("\"title\":" + JsonSerializer.Serialize(title), "<text variable=\"title\" quotes=\"true\" suffix=\".\"/>"));

    [Fact]
    public void EmptyLocaleQuoteTermsDoNotRequireTextBoundaryNodes() =>
        Assert.Equal("<i>Title</i>", Render("\"title\":\"Title\"", "<text variable=\"title\" quotes=\"true\" font-style=\"italic\"/>",
            locale: "<locale><terms><term name=\"open-quote\"></term><term name=\"close-quote\"></term></terms></locale>",
            options: new CslRenderOptions { OutputFormat = CslOutputFormat.Html }));

    [Theory]
    [InlineData("One <i>Two <i>Three</i> Four</i> Five", "font-style=\"italic\"", "<i>One <span style=\"font-style:normal;\">Two <i>Three</i> Four</span> Five</i>")]
    [InlineData("His <sc>Anonymous</sc> Life", "font-variant=\"small-caps\"", "<span style=\"font-variant:small-caps;\">His <span style=\"font-variant:normal;\">Anonymous</span> Life</span>")]
    [InlineData("Lessard <span class=\"nodecor\">v.</span> Schmidt", "font-style=\"italic\" text-case=\"capitalize-all\"", "<i>Lessard <span style=\"font-style:normal;\">v.</span> Schmidt</i>")]
    [InlineData("DNA and <span class=\"nocase\">iPhone</span>", "text-case=\"sentence\"", "DNA and iPhone")]
    public void RichTextEmphasisAndCaseProtectionRespectEnclosingStyle(string value, string formatting, string expected) =>
        Assert.Equal(expected, Render("\"title\":" + JsonSerializer.Serialize(value), "<text variable=\"title\" " + formatting + "/>", options: new CslRenderOptions { OutputFormat = CslOutputFormat.Html }));

    [Theory]
    [InlineData(CslOutputFormat.PlainText)]
    [InlineData(CslOutputFormat.Html)]
    public void DeepRichTextIsRejectedBeforeRecursiveFormatting(CslOutputFormat format) {
        string value = string.Concat(Enumerable.Repeat("<i>", 80)) + "deep" + string.Concat(Enumerable.Repeat("</i>", 80));
        Assert.Throws<InvalidDataException>(() => Render("\"title\":" + JsonSerializer.Serialize(value), "<text variable=\"title\"/>", options: new CslRenderOptions { OutputFormat = format }));
    }

    [Fact]
    public void RichTextEscapesUnsupportedTagsAndDropsExecutableAttributes() {
        string value = "<i onclick=\"run()\">safe</i><script>bad()</script>";
        string result = Render("\"title\":" + JsonSerializer.Serialize(value), "<text variable=\"title\"/>", options: new CslRenderOptions { OutputFormat = CslOutputFormat.Html });
        Assert.Contains("<i>safe</i>", result);
        Assert.DoesNotContain("onclick", result);
        Assert.DoesNotContain("<script>", result);
        Assert.Contains("&lt;script&gt;", result);
    }
    [Theory]
    [InlineData("<text value=\"x\" font-style=\"italic; background:url(https://example.org)\"/>")]
    [InlineData("<foreign:text xmlns:foreign=\"urn:vendor\" value=\"x\"/>")]
    public void InvalidPresentationAndForeignRenderingVocabularyIsRejected(string layout) =>
        Assert.Throws<InvalidDataException>(() => CslStyle.Parse(Header + "<citation><layout>" + layout + "</layout></citation></style>"));
    private static string Render(string fields, string layout, string styleOptions = "", string locale = "", CslRenderOptions? options = null) {
        BibliographyDocument data = BibliographyDocument.Parse("[{\"id\":\"a\",\"type\":\"book\"," + fields + "}]", BibliographyFormat.CslJson).Document;
        CslStyle style = CslStyle.Parse(Header.Replace("class=\"in-text\"", "class=\"in-text\" " + styleOptions) + locale + "<citation><layout>" + layout + "</layout></citation></style>");
        var cite = new CslCitation("cluster"); cite.Items.Add(new CslCitationItem("a"));
        return new CslProcessor(data, style, options).Render(new[] { cite }).Citations.Single().Content;
    }

    [Theory]
    [InlineData("42", "42nd")]
    [InlineData("112", "112th")]
    [InlineData("121", "121st")]
    public void OrdinalsUseTheDefaultLastDigitAndTeenRules(string value, string expected) =>
        Assert.Equal(expected, Render("\"volume\":\"" + value + "\"", "<number variable=\"volume\" form=\"ordinal\"/>"));

    [Fact]
    public void ALocalOrdinalDefinitionReplacesTheInheritedSuffixSet() =>
        Assert.Equal("42x", Render("\"volume\":\"42\"", "<number variable=\"volume\" form=\"ordinal\"/>", locale: "<locale><terms><term name=\"ordinal\">x</term></terms></locale>"));

    [Theory]
    [InlineData("M.E.", ". ", true, "M. E. Doe")]
    [InlineData("John M.E.", "", true, "JME Doe")]
    [InlineData("John M.E.", ". ", false, "John M. E. Doe")]
    [InlineData("James T Kirk", ".", false, "James T. Kirk Doe")]
    [InlineData("John Bertrand de Cusance Morant", ".", true, "J.B. de C.M. Doe")]
    public void ExistingInitialsAreNormalizedWithoutLosingParts(string given, string suffix, bool initialize, string expected) =>
        Assert.Equal(expected, Render("\"author\":[{\"family\":\"Doe\",\"given\":" + JsonSerializer.Serialize(given) + "}]", "<names variable=\"author\"><name initialize-with=\"" + suffix + "\" initialize=\"" + initialize.ToString().ToLowerInvariant() + "\"/></names>"));

    [Fact]
    public void NameAffixesApplyOnceAndExplicitLabelPluralityIsHonored() =>
        Assert.Equal("[Doe and Roe] ed.", Render("\"editor\":[{\"family\":\"Doe\"},{\"family\":\"Roe\"}]", "<names variable=\"editor\"><name prefix=\"[\" suffix=\"]\" and=\"text\"/><label form=\"short\" plural=\"never\" prefix=\" \"/></names>"));

    [Theory]
    [InlineData("L2d", "yes|L2d")]
    [InlineData("2nd", "yes|2nd")]
    [InlineData("2nd edition", "no|2nd edition")]
    [InlineData("2 ,3& 4", "yes|2nd, 3rd &amp; 4th")]
    public void NumericConditionsAndFormattingRespectAffixesAndSeparators(string value, string expected) =>
        Assert.Equal(expected, Render("\"volume\":" + JsonSerializer.Serialize(value), "<choose><if is-numeric=\"volume\"><text value=\"yes\"/></if><else><text value=\"no\"/></else></choose><number variable=\"volume\" form=\"ordinal\" prefix=\"|\"/>", options: new CslRenderOptions { OutputFormat = CslOutputFormat.Html }));

    [Fact]
    public void AlphabeticPagesAndCountsHaveCorrectLabels() =>
        Assert.Equal("pages S213–S235; 0 page", Render("\"page\":\"S213-S235\",\"number-of-pages\":0", "<group delimiter=\"; \"><group delimiter=\" \"><label variable=\"page\"/><text variable=\"page\"/></group><group delimiter=\" \"><text variable=\"number-of-pages\"/><label variable=\"number-of-pages\"/></group></group>", "page-range-format=\"expanded\""));

    [Fact]
    public void PlainTextOutputLimitsDoNotCountHtmlEscapingOrInvisibleMarkup() {
        var options = new CslRenderOptions { MaximumOutputCharacters = 1 };
        Assert.Equal("&", Render("\"title\":\"&amp;\"", "<text variable=\"title\" font-style=\"italic\"/>", options: options));
        options.MaximumIntermediateCharacters = 5;
        Assert.Throws<InvalidDataException>(() => Render("\"title\":\"&amp;\"", "<text variable=\"title\" font-style=\"italic\"/>", options: options));
    }

    [Fact]
    public void EntityEscapingDoesNotReduceTheAcceptedDecodedFieldBudget() {
        string title = new string('&', 850000);
        Assert.Equal(title, Render("\"title\":" + JsonSerializer.Serialize(title), "<text variable=\"title\"/>"));
    }

    [Fact]
    public void ProcessorsExposeAndCanRejectApproximateCitationData() {
        BibliographyDocument data = BibliographyDocument.Parse("@book{x,title={A {DNA} study}}", BibliographyFormat.BibLatex).Document;
        CslStyle style = CslStyle.Parse(Header + "<citation><layout><text variable=\"title\"/></layout></citation></style>");
        Assert.Contains(new CslProcessor(data, style).DataConversionReport.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV252");
        Assert.Throws<BibliographyConversionLossException>(() => new CslProcessor(data, style, new CslRenderOptions { RequireNoDataLoss = true }));
        Assert.Equal("A {DNA} study", data.Items.Single().Title);
    }
}
