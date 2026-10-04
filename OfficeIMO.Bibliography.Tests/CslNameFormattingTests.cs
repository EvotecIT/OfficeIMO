using System.Text.Json;

namespace OfficeIMO.Bibliography.Tests;

public sealed class CslNameFormattingTests {
    [Theory]
    [InlineData("<b>John</b> Quiggly", "true", "<b>Doe</b>, <b>J.</b> Q.")]
    [InlineData("<b>John</b>-Quiggly", "true", "<b>Doe</b>, <b>J.</b>-Q.")]
    [InlineData("<b>J.</b> Q.", "false", "<b>Doe</b>, <b>J.</b> Q.")]
    [InlineData("<b>J.</b>-Q.", "false", "<b>Doe</b>, <b>J.</b>-Q.")]
    [InlineData("<b>José</b> André", "true", "<b>Doe</b>, <b>J.</b> A.")]
    [InlineData("<b>J<i>o</i>hn</b> Quiggly", "true", "<b>Doe</b>, <b>J.</b> Q.")]
    public void InitialsPreserveEmphasisAndNormalizeOnlyVisibleNameText(string given, string initialize, string expected) =>
        Assert.Equal(expected, Render("{\"family\":\"<b>Doe</b>\",\"given\":" + JsonSerializer.Serialize(given) + "}",
            "<name name-as-sort-order=\"all\" initialize-with=\". \" initialize=\"" + initialize + "\"/>", html: true));
    [Theory]
    [InlineData("d'", "John d’Jones")]
    [InlineData("d’", "John d’Jones")]
    [InlineData("al-", "John al-Jones")]
    public void DroppingParticlesJoinSurnamesWithoutInventedSpaces(string particle, string expected) =>
        Assert.Equal(expected, Render("{\"family\":\"Jones\",\"given\":\"John\",\"dropping-particle\":" + JsonSerializer.Serialize(particle) + "}", "<name/>"));

    [Theory]
    [InlineData("", "display-and-sort", "[Jean] (de La Fontaine III)")]
    [InlineData("name-as-sort-order=\"all\" sort-separator=\" \"", "sort-only", "(La Fontaine) [Jean de] III")]
    [InlineData("name-as-sort-order=\"all\" sort-separator=\" \"", "display-and-sort", "(Fontaine) [Jean de La] III")]
    public void NamePartAffixesEncloseParticlesAndSuffixesInTheirDisplayPositions(string nameOptions, string demote, string expected) =>
        Assert.Equal(expected, Render("{\"family\":\"Fontaine\",\"given\":\"Jean\",\"dropping-particle\":\"de\",\"non-dropping-particle\":\"La\",\"suffix\":\"III\"}",
            "<name " + nameOptions + "><name-part name=\"family\" prefix=\"(\" suffix=\")\"/><name-part name=\"given\" prefix=\"[\" suffix=\"]\"/></name>", "demote-non-dropping-particle=\"" + demote + "\""));

    [Fact]
    public void NamePartFormattingDoesNotDecorateTheSuffixAndUsesGivenFormattingForDroppingParticles() =>
        Assert.Equal("<i>Jean</i> (<i>de</i> <b>La Fontaine</b> III)", Render(
            "{\"family\":\"Fontaine\",\"given\":\"Jean\",\"dropping-particle\":\"de\",\"non-dropping-particle\":\"La\",\"suffix\":\"III\"}",
            "<name><name-part name=\"family\" font-weight=\"bold\" prefix=\"(\" suffix=\")\"/><name-part name=\"given\" font-style=\"italic\"/></name>", html: true));

    [Fact]
    public void InstitutionalNamesHonorFamilyFormattingAndAreNotInvertedForDelimiterRules() {
        const string names = "{\"literal\":\"Agency one\"},{\"literal\":\"Agency two\"}";
        Assert.Equal("AGENCY ONE & AGENCY TWO", Render(names,
            "<name name-as-sort-order=\"all\" and=\"symbol\" delimiter-precedes-last=\"after-inverted-name\"><name-part name=\"family\" text-case=\"uppercase\"/></name>"));
        Assert.Equal("JONES, JOHN, & DOE, JANE", Render("{\"family\":\"Jones\",\"given\":\"John\"},{\"family\":\"Doe\",\"given\":\"Jane\"}",
            "<name name-as-sort-order=\"all\" and=\"symbol\" delimiter-precedes-last=\"after-inverted-name\" text-case=\"uppercase\"/>"));
    }

    [Fact]
    public void SoleGivenNamesAndCallerNonbreakingSpacingRemainVisible() {
        Assert.Equal("Banksy", Render("{\"given\":\"Banksy\"}", "<name initialize-with=\".\"/>"));
        Assert.Equal("John\u00a0Doe", Render("{\"given\":\"John\u00a0\",\"family\":\"Doe\"}", "<name/>"));
        Assert.Equal("van Gogh", Render("{\"given\":\"Vincent\",\"family\":\"Gogh\",\"non-dropping-particle\":\"van\"}", "<name form=\"short\" name-as-sort-order=\"all\"/>"));
    }

    private static string Render(string names, string formatting, string styleOptions = "", bool html = false) {
        BibliographyDocument data = BibliographyDocument.Parse("[{\"id\":\"a\",\"type\":\"book\",\"author\":[" + names + "]}]", BibliographyFormat.CslJson).Document;
        CslStyle style = CslStyle.Parse("<style xmlns=\"http://purl.org/net/xbiblio/csl\" version=\"1.0\" class=\"in-text\" " + styleOptions + "><citation><layout><names variable=\"author\">" + formatting + "</names></layout></citation></style>");
        var citation = new CslCitation("cluster"); citation.Items.Add(new CslCitationItem("a"));
        return new CslProcessor(data, style, new CslRenderOptions { OutputFormat = html ? CslOutputFormat.Html : CslOutputFormat.PlainText }).Render(new[] { citation }).Citations.Single().Content;
    }
}
