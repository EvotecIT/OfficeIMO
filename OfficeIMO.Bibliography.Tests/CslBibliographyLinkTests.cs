using System.Text.Json;

namespace OfficeIMO.Bibliography.Tests;

public sealed class CslBibliographyLinkTests {
    [Theory]
    [InlineData("URL", "https://example.org/?a=1&b=2", "https://example.org/?a=1&amp;b=2")]
    [InlineData("DOI", "10.1234/example", "https://doi.org/10.1234/example")]
    [InlineData("PMID", "12345", "https://www.ncbi.nlm.nih.gov/pubmed/12345")]
    [InlineData("PMCID", "PMC12345", "https://www.ncbi.nlm.nih.gov/pmc/articles/PMC12345")]
    public void BibliographyIdentifiersBecomeLinksWhileCitationIdentifiersStayText(string variable, string value, string target) {
        CslRenderResult result = Render(variable, value);
        Assert.Equal(System.Net.WebUtility.HtmlEncode(value), result.Citations.Single().Content);
        Assert.Equal("<div class=\"csl-entry\"><a href=\"" + target + "\">" + System.Net.WebUtility.HtmlEncode(value) + "</a></div>", result.Bibliography.Single().Content);
    }

    [Theory]
    [InlineData("DOI", "10.1234/example", "https://doi.org/")]
    [InlineData("PMID", "12345", "https://www.ncbi.nlm.nih.gov/pubmed/")]
    [InlineData("PMCID", "PMC12345", "https://www.ncbi.nlm.nih.gov/pmc/articles/")]
    public void UriPrefixesJoinTheAnchorAndOtherAffixesRemainOutsideIt(string variable, string value, string uriPrefix) {
        string layout = "<group prefix=\"Available from: \"><text variable=\"" + variable + "\" prefix=\"" + uriPrefix + "\" suffix=\".\" font-style=\"italic\"/></group>";
        string linked = Render(variable, value, layout).Bibliography.Single().Content;
        Assert.Equal("<div class=\"csl-entry\">Available from: <i><a href=\"" + uriPrefix + value + "\">" + uriPrefix + value + "</a></i>.</div>", linked);
        Assert.Equal("Available from: " + uriPrefix + value + ".", Render(variable, value, layout, new CslRenderOptions()).Bibliography.Single().Content);
    }

    [Theory]
    [InlineData("javascript:alert(1)")]
    [InlineData("data:text/html,hello")]
    [InlineData("file:///C:/secret.txt")]
    [InlineData("../reference.html")]
    [InlineData("https://example.org/\nreference")]
    public void UnsafeOrNonWebUrlsRemainVisibleWithoutAnAnchor(string value) {
        string output = Render("URL", value).Bibliography.Single().Content;
        Assert.DoesNotContain("<a ", output);
        Assert.Contains(System.Net.WebUtility.HtmlEncode(value), output);
    }

    [Fact]
    public void LinkingCanBeDisabledAndTheProcessorOwnsItsOptionsSnapshot() {
        var options = new CslRenderOptions { OutputFormat = CslOutputFormat.Html, LinkBibliographyIdentifiers = false };
        CslProcessor processor = Processor("DOI", "10.1234/example", null, options);
        options.LinkBibliographyIdentifiers = true;
        Assert.Equal("<div class=\"csl-entry\">10.1234/example</div>", processor.RenderBibliography().Single().Content);
    }

    [Fact]
    public void LinkMarkupCountsTowardTheIntermediateOutputBudget() =>
        Assert.Throws<InvalidDataException>(() => Render("DOI", "10.1234/example", options: new CslRenderOptions {
            OutputFormat = CslOutputFormat.Html, MaximumIntermediateCharacters = 30
        }));

    private static CslRenderResult Render(string variable, string value, string? layout = null, CslRenderOptions? options = null) {
        var citation = new CslCitation("cite"); citation.Items.Add(new CslCitationItem("one"));
        return Processor(variable, value, layout, options ?? new CslRenderOptions { OutputFormat = CslOutputFormat.Html }).Render(new[] { citation });
    }

    private static CslProcessor Processor(string variable, string value, string? layout, CslRenderOptions options) {
        BibliographyDocument document = BibliographyDocument.Parse("[{\"id\":\"one\",\"type\":\"book\",\"" + variable + "\":" + JsonSerializer.Serialize(value) + "}]", BibliographyFormat.CslJson).Document;
        layout ??= "<text variable=\"" + variable + "\"/>";
        CslStyle style = CslStyle.Parse("<style xmlns=\"http://purl.org/net/xbiblio/csl\" version=\"1.0\" class=\"in-text\"><citation><layout><text variable=\"" + variable + "\"/></layout></citation><bibliography><layout>" + layout + "</layout></bibliography></style>");
        return new CslProcessor(document, style, options);
    }
}
