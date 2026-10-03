using System.Text.Json;

namespace OfficeIMO.Bibliography.Tests;

public sealed class CslPunctuationBoundaryTests {
    [Theory]
    [InlineData("Done:", ": next", "Done: next")]
    [InlineData("Done;", ". next", "Done; next")]
    [InlineData("Done!", ": next", "Done! next")]
    [InlineData("Done,", ", next", "Done, next")]
    [InlineData("Done:", "! next", "Done! next")]
    [InlineData("Done;", "? next", "Done? next")]
    [InlineData("Done.", "; next", "Done.; next")]
    [InlineData("Done,", ". next", "Done,. next")]
    public void FieldBoundariesResolveCollisionsWithoutChangingIntentionalPairs(string left, string right, string expected) {
        Assert.Equal(expected, Render(left, right, "<text variable=\"title\"/><text variable=\"abstract\"/>"));
        Assert.Equal(expected, Render(left, "next", "<group delimiter=\"" + right.Substring(0, right.Length - 4) + "\"><text variable=\"title\"/><text variable=\"abstract\"/></group>"));
        Assert.Equal(expected, Render(left, "", "<text variable=\"title\" suffix=\"" + right + "\"/>"));
    }

    [Fact]
    public void PunctuationCleanupRetainsFormattingAndOnlyTouchesFieldBoundaries() {
        Assert.Equal("<i>Done</i><b>! next</b>", Render("Done:", "! next", "<text variable=\"title\" font-style=\"italic\"/><text variable=\"abstract\" font-weight=\"bold\"/>", true));
        Assert.Equal("<i>Done,</i><b> next</b>", Render("Done,", ", next", "<text variable=\"title\" font-style=\"italic\"/><text variable=\"abstract\" font-weight=\"bold\"/>", true));
        Assert.Equal("Keep::this??text", Render("Keep::this??text", "", "<text variable=\"title\"/>"));
    }

    private static string Render(string title, string body, string layout, bool html = false) {
        string json = JsonSerializer.Serialize(new[] { new { id = "one", type = "book", title, @abstract = body } });
        BibliographyDocument document = BibliographyDocument.Parse(json, BibliographyFormat.CslJson).Document;
        CslStyle style = CslStyle.Parse("<style xmlns=\"http://purl.org/net/xbiblio/csl\" version=\"1.0\" class=\"in-text\"><citation><layout>" + layout + "</layout></citation></style>");
        var citation = new CslCitation("cite");
        citation.Items.Add(new CslCitationItem("one"));
        return new CslProcessor(document, style, new CslRenderOptions { OutputFormat = html ? CslOutputFormat.Html : CslOutputFormat.PlainText }).Render(new[] { citation }).Citations.Single().Content;
    }
}
