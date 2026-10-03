using System.Text.Json;

namespace OfficeIMO.Bibliography.Tests;

public sealed class CslCaseContractTests {
    [Theory]
    [InlineData("Review of Book by A.N. Author", "title", "Review of Book by A.N. Author")]
    [InlineData("Some rules according to the judge", "title", "Some Rules according to the Judge")]
    [InlineData("A view from 2016 with significant utility", "title", "A View from 2016 with Significant Utility")]
    [InlineData("the evidence: according to a report", "title", "The Evidence: According to a Report")]
    [InlineData("two-thirds of the sample", "title", "Two-Thirds of the Sample")]
    [InlineData("Employee pro-environmental behavior", "title", "Employee pro-Environmental Behavior")]
    [InlineData("response to Shafi`i comment", "title", "Response to Shafi`i Comment")]
    [InlineData("TITLE WITH DATA", "title", "Title with Data")]
    [InlineData("an eBay and iPhone story", "capitalize-all", "An eBay And iPhone Story")]
    [InlineData("eBay title", "capitalize-first", "eBay title")]
    [InlineData("a title", "capitalize-first", "A title")]
    [InlineData("ALL CAPS TITLE", "sentence", "All caps title")]
    [InlineData("a mixed Title", "sentence", "A mixed Title")]
    public void CasingFollowsWordsInitialsPhrasesAndTheRequestedMode(string title, string mode, string expected) =>
        Assert.Equal(expected, Render(title, mode));

    [Fact]
    public void CaseProtectedWordsAndInputEmphasisKeepTheirOwnedFormatting() =>
        Assert.Equal("A <i>Study</i> with DNA", Render("a <i>study</i> with <span class=\"nocase\">DNA</span>", "title", format: CslOutputFormat.Html));

    [Fact]
    public void ItemLanguageSelectsUnicodeCasing() =>
        Assert.Equal("İSTANBUL", Render("istanbul", "uppercase", language: "tr"));

    [Theory]
    [InlineData("en-US", "fr-FR", "", "A Study with Results")]
    [InlineData("fr-FR", "en-US", "", "a study with results")]
    [InlineData("fr-FR", "fr-FR", "en-GB", "A Study with Results")]
    public void TitleCaseUsesStyleAndItemLanguageIndependentlyOfOutputLocale(string styleLanguage, string outputLanguage, string itemLanguage, string expected) =>
        Assert.Equal(expected, Render("a study with results", "title", itemLanguage, styleLanguage, outputLanguage));

    private static string Render(string title, string mode, string language = "", string styleLanguage = "en-US", string? outputLanguage = null, CslOutputFormat format = CslOutputFormat.PlainText) {
        BibliographyDocument document = BibliographyDocument.Parse("[{\"id\":\"one\",\"type\":\"book\",\"language\":" + JsonSerializer.Serialize(language) + ",\"title\":" + JsonSerializer.Serialize(title) + "}]", BibliographyFormat.CslJson).Document;
        CslStyle style = CslStyle.Parse("<style xmlns=\"http://purl.org/net/xbiblio/csl\" version=\"1.0\" class=\"in-text\" default-locale=\"" + styleLanguage + "\"><citation><layout><text variable=\"title\" text-case=\"" + mode + "\"/></layout></citation></style>");
        var cite = new CslCitation("cite"); cite.Items.Add(new CslCitationItem("one"));
        return new CslProcessor(document, style, new CslRenderOptions { Locale = outputLanguage, OutputFormat = format }).Render(new[] { cite }).Citations.Single().Content;
    }
}
