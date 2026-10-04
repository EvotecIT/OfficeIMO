using System.Text.Json;

namespace OfficeIMO.Bibliography.Tests;

public sealed class CslQuoteBoundaryTests {
    [Theory]
    [InlineData("Title", ", Next", true, "“Title,” Next", "“<i>Title</i><b>,</b>”<b> Next</b>")]
    [InlineData("Title", ", Next", false, "“Title”, Next", "“<i>Title</i>”<b>, Next</b>")]
    [InlineData("Title", ". Next", true, "“Title.” Next", "“<i>Title</i><b>.</b>”<b> Next</b>")]
    [InlineData("Title?", ". Next", true, "“Title?” Next", "“<i>Title?</i>”<b> Next</b>")]
    [InlineData("Title", ",. Next", true, "“Title,.” Next", "“<i>Title</i><b>,.</b>”<b> Next</b>")]
    public void QuotedFieldsLocalizePunctuationAcrossFormatting(string title, string body, bool inside, string plain, string html) {
        const string fields = "<text variable=\"title\" quotes=\"true\" font-style=\"italic\"/><text variable=\"abstract\" font-weight=\"bold\"/>";
        Assert.Equal(plain, Render(title, body, fields, inside));
        Assert.Equal(html, Render(title, body, fields, inside, CslOutputFormat.Html));
    }

    [Theory]
    [InlineData("Singin’", false, "Singin’, Next")]
    [InlineData("’90", false, "’90, Next")]
    [InlineData("Singin’", true, "“Singin’,” Next")]
    public void ApostrophesDoNotActAsClosingQuotationMarks(string title, bool quotes, string expected) {
        string fields = "<text variable=\"title\" quotes=\"" + (quotes ? "true" : "false") + "\" suffix=\", \"/><text variable=\"abstract\"/>";
        Assert.Equal(expected, Render(title, "Next", fields));
        Assert.Equal(expected, Render(title, "Next", fields, format: CslOutputFormat.Html));
    }

    [Fact]
    public void InputQuotesRetainTheirOwnFormattingWhenPunctuationMoves() {
        const string fields = "<text variable=\"title\"/><text variable=\"abstract\" font-weight=\"bold\"/>";
        Assert.Equal("“Title,” Next", Render("<i>“Title”</i>", ", Next", fields));
        Assert.Equal("<i>“Title</i><b>,</b><i>”</i><b> Next</b>", Render("<i>“Title”</i>", ", Next", fields, format: CslOutputFormat.Html));
    }

    [Theory]
    [InlineData(CslOutputFormat.PlainText)]
    [InlineData(CslOutputFormat.Html)]
    public void AdjacentPunctuationMovesAsOneRunAcrossCustomQuoteTerms(CslOutputFormat format) {
        string locale = "<terms><term name=\"open-quote\">[[</term><term name=\"close-quote\">]]</term></terms>";
        const string fields = "<text variable=\"title\" quotes=\"true\" suffix=\",\"/><text variable=\"abstract\" prefix=\". \"/>";
        Assert.Equal("[[Title,.]] Next", Render("Title", "Next", fields, format: format, locale: locale));
    }

    [Fact]
    public void MovingPunctuationRetainsInheritedFormattingAndInputFlips() {
        const string fields = "<group font-style=\"italic\"><text variable=\"title\" quotes=\"true\"/><text variable=\"abstract\" font-weight=\"bold\"/></group>";
        Assert.Equal("<i>“<span style=\"font-style:normal;\">Title</span><b>,</b>”<b> Next</b></i>",
            Render("<i>Title</i>", ", Next", fields, format: CslOutputFormat.Html));
    }

    [Fact]
    public void InputCannotDeclareAnApostropheToBeAClosingQuote() {
        Assert.Equal("Singin’, Next", Render("Singin<span data-csl-quote-end=\"true\">’</span>", ", Next",
            "<text variable=\"title\"/><text variable=\"abstract\"/>", format: CslOutputFormat.Html));
    }

    [Theory]
    [InlineData("\"ETFA '09\"", "“ETFA ’09”")]
    [InlineData("\"ETFA '09 and '10\"", "“ETFA ’09 and ’10”")]
    [InlineData("'His \"orphan title'", "“His \"orphan title”")]
    [InlineData("\"ETFA '09 and 'nested title'\"", "“ETFA ’09 and ‘nested title’”")]
    public void UnmatchedInnerCandidatesDoNotBlockACompleteOuterPair(string title, string expected) {
        const string fields = "<text variable=\"title\"/>";
        Assert.Equal(expected, Render(title, string.Empty, fields));
        Assert.Equal(System.Net.WebUtility.HtmlEncode(expected), Render(title, string.Empty, fields, format: CslOutputFormat.Html));
    }

    [Fact]
    public void AbbreviatedYearsInsideInputQuotesRetainTheirFormattingAndQuoteProvenance() {
        const string fields = "<text variable=\"title\"/><text variable=\"abstract\" font-weight=\"bold\"/>";
        Assert.Equal("<i>“ETFA ’09</i><b>,</b><i>”</i><b> Next</b>", Render("<i>\"ETFA '09\"</i>", ", Next", fields, format: CslOutputFormat.Html));
    }

    [Fact]
    public void BibliographyLinksStayOnTheIdentifierWhenPunctuationMoves() {
        var document = BibliographyDocument.Parse("[{\"id\":\"one\",\"type\":\"book\",\"URL\":\"https://example.org/\",\"abstract\":\", Next\"}]", BibliographyFormat.CslJson).Document;
        var style = CslStyle.Parse("<style xmlns=\"http://purl.org/net/xbiblio/csl\" version=\"1.0\" class=\"in-text\"><citation><layout><text variable=\"id\"/></layout></citation><bibliography><layout><text variable=\"URL\" quotes=\"true\"/><text variable=\"abstract\" font-weight=\"bold\"/></layout></bibliography></style>");
        Assert.Equal("<div class=\"csl-entry\">“<a href=\"https://example.org/\">https://example.org/</a><b>,</b>”<b> Next</b></div>",
            new CslProcessor(document, style, new CslRenderOptions { OutputFormat = CslOutputFormat.Html }).RenderBibliography().Single().Content);
    }

    [Fact]
    public void PunctuationDoesNotMoveBetweenDisplayContainers() {
        var document = BibliographyDocument.Parse("[{\"id\":\"one\",\"type\":\"book\",\"title\":\"Title\",\"abstract\":\", Next\"}]", BibliographyFormat.CslJson).Document;
        var style = CslStyle.Parse("<style xmlns=\"http://purl.org/net/xbiblio/csl\" version=\"1.0\" class=\"in-text\"><citation><layout><text variable=\"id\"/></layout></citation><bibliography><layout><text variable=\"title\" quotes=\"true\" display=\"block\"/><text variable=\"abstract\" display=\"block\"/></layout></bibliography></style>");
        Assert.Equal("<div class=\"csl-entry\"><div class=\"csl-block\">“Title”</div><div class=\"csl-block\">, Next</div></div>",
            new CslProcessor(document, style, new CslRenderOptions { OutputFormat = CslOutputFormat.Html }).RenderBibliography().Single().Content);
    }

    private static string Render(string title, string body, string fields, bool inside = true, CslOutputFormat format = CslOutputFormat.PlainText, string locale = "") {
        string json = JsonSerializer.Serialize(new[] { new { id = "one", type = "book", title, @abstract = body } });
        var document = BibliographyDocument.Parse(json, BibliographyFormat.CslJson).Document;
        var style = CslStyle.Parse("<style xmlns=\"http://purl.org/net/xbiblio/csl\" version=\"1.0\" class=\"in-text\"><locale><style-options punctuation-in-quote=\"" + (inside ? "true" : "false") + "\"/>" + locale + "</locale><citation><layout>" + fields + "</layout></citation></style>");
        var citation = new CslCitation("one"); citation.Items.Add(new CslCitationItem("one"));
        return new CslProcessor(document, style, new CslRenderOptions { OutputFormat = format }).Render(new[] { citation }).Citations.Single().Content;
    }
}
