using System.Text.Json;

namespace OfficeIMO.Bibliography.Tests;

public sealed class CslAdjacentQuoteContractTests {
    [Theory]
    [InlineData("'\"Title\"'", false, "“‘Title.’”", "“‘Title.’”")]
    [InlineData("\"'Title'\"", false, "“‘Title.’”", "“‘Title.’”")]
    [InlineData("<i>'\"Title\"'</i>", false, "“‘Title.’”", "<i>“‘Title</i>.<i>’”</i>")]
    [InlineData("<i>'</i><b>\"Title\"</b><i>'</i>", false, "“‘Title.’”", "<i>“</i><b>‘Title</b>.<b>’</b><i>”</i>")]
    [InlineData("'\"Title\"'", true, "“‘“Title.”’”", "“‘“Title.”’”")]
    [InlineData("\"'Title'\"", true, "“‘“Title.”’”", "“‘“Title.”’”")]
    [InlineData("Before '\"Title\"' after", false, "Before “‘Title’” after.", "Before “‘Title’” after.")]
    [InlineData("'\"Title", false, "’\"Title.", "’&quot;Title.")]
    public void AdjacentMixedQuotesRetainQuotationLevelsAndFormatting(string title, bool styleQuotes, string plain, string html) {
        CslStyle style = CslStyle.Parse(Header + "<citation><layout><text variable=\"title\" quotes=\"" + styleQuotes.ToString().ToLowerInvariant() + "\" suffix=\".\"/></layout></citation></style>");
        Assert.Equal(plain, Render(title, style, CslOutputFormat.PlainText));
        Assert.Equal(html, Render(title, style, CslOutputFormat.Html));
    }

    [Theory]
    [InlineData(CslOutputFormat.PlainText)]
    [InlineData(CslOutputFormat.Html)]
    public void AdjacentMixedQuotesUseBothCallerLocaleQuotationLevels(CslOutputFormat format) {
        CslStyle style = CslStyle.Parse(Header + "<locale><terms><term name=\"open-quote\">[[</term><term name=\"close-quote\">]]</term><term name=\"open-inner-quote\">{{</term><term name=\"close-inner-quote\">}}</term></terms></locale><citation><layout><text variable=\"title\" suffix=\".\"/></layout></citation></style>");
        Assert.Equal("[[{{Title.}}]]", Render("'\"Title\"'", style, format));
    }

    private const string Header = "<style xmlns=\"http://purl.org/net/xbiblio/csl\" version=\"1.0\" class=\"in-text\">";

    private static string Render(string title, CslStyle style, CslOutputFormat format) {
        BibliographyDocument data = BibliographyDocument.Parse(JsonSerializer.Serialize(new[] { new { id = "one", type = "book", title } }), BibliographyFormat.CslJson).Document;
        var cite = new CslCitation("cite");
        cite.Items.Add(new CslCitationItem("one"));
        return new CslProcessor(data, style, new CslRenderOptions { OutputFormat = format }).Render(new[] { cite }).Citations.Single().Content;
    }
}
