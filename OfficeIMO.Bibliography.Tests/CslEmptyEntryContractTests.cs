namespace OfficeIMO.Bibliography.Tests;

public sealed class CslEmptyEntryContractTests {
    private const string Header = "<style xmlns=\"http://purl.org/net/xbiblio/csl\" version=\"1.0\" class=\"in-text\">";

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void OmittedBibliographyTextRetainsItsKeyAndCitationNumber(bool html, bool bibliographyNumbers) {
        const string json = """
            [{"id":"a","type":"book","title":"Alpha"},
             {"id":"b","type":"personal_communication","title":"Beta"},
             {"id":"c","type":"book","title":"Gamma"}]
            """;
        string number = bibliographyNumbers ? "<text variable=\"citation-number\" suffix=\". \"/>" : "";
        CslStyle style = CslStyle.Parse(Header + "<citation><layout delimiter=\";\"><text variable=\"citation-number\"/></layout></citation><bibliography><sort><key variable=\"title\"/></sort><layout><choose><if type=\"book\">" + number + "<text variable=\"title\"/></if></choose></layout></bibliography></style>");
        var cluster = new CslCitation("cluster"); foreach (string key in new[] { "a", "b", "c" }) cluster.Items.Add(new CslCitationItem(key));
        CslRenderResult result = new CslProcessor(BibliographyDocument.Parse(json, BibliographyFormat.CslJson).Document, style, new CslRenderOptions { OutputFormat = html ? CslOutputFormat.Html : CslOutputFormat.PlainText }).Render(new[] { cluster });
        Assert.Equal("1;2;3", result.Citations.Single().Content);
        Assert.Equal(new[] { "a", "b", "c" }, result.Bibliography.Select(entry => entry.Key));
        Assert.Equal(new[] { false, true, false }, result.Bibliography.Select(entry => entry.IsEmpty));
        Assert.Equal(html ? "<div class=\"csl-entry\"></div>" : "", result.Bibliography[1].Content);
        Assert.Equal(new[] { "a", "c" }, result.Bibliography.Where(entry => !entry.IsEmpty).Select(entry => entry.Key));
        Assert.Contains(bibliographyNumbers ? "3. Gamma" : "Gamma", result.Bibliography[2].Content);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void APrintedCitationNumberIsContentEvenWhenTheTitleIsMissing(bool html) {
        BibliographyDocument data = BibliographyDocument.Parse("[{\"id\":\"a\",\"type\":\"book\"}]", BibliographyFormat.CslJson).Document;
        CslStyle style = CslStyle.Parse(Header + "<citation><layout><text variable=\"title\"/></layout></citation><bibliography><layout><text variable=\"citation-number\" prefix=\"[\" suffix=\"]\"/><text variable=\"title\"/></layout></bibliography></style>");
        CslRenderedEntry entry = new CslProcessor(data, style, new CslRenderOptions { OutputFormat = html ? CslOutputFormat.Html : CslOutputFormat.PlainText }).RenderBibliography().Single();
        Assert.False(entry.IsEmpty);
        Assert.Equal(html ? "<div class=\"csl-entry\">[1]</div>" : "[1]", entry.Content);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ExplicitStyleWhitespaceIsRetainedAsContent(bool html) {
        BibliographyDocument data = BibliographyDocument.Parse("[{\"id\":\"a\",\"type\":\"book\"}]", BibliographyFormat.CslJson).Document;
        CslStyle style = CslStyle.Parse(Header + "<citation><layout><text value=\" \"/></layout></citation><bibliography><layout><text value=\" \"/></layout></bibliography></style>");
        CslRenderedEntry entry = new CslProcessor(data, style, new CslRenderOptions { OutputFormat = html ? CslOutputFormat.Html : CslOutputFormat.PlainText }).RenderBibliography().Single();
        Assert.False(entry.IsEmpty);
        Assert.Equal(html ? "<div class=\"csl-entry\"> </div>" : " ", entry.Content);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void EmptinessReflectsTheTextAfterLocaleQuoteNormalization(bool html) {
        BibliographyDocument data = BibliographyDocument.Parse("[{\"id\":\"a\",\"type\":\"book\",\"title\":\"“”\"}]", BibliographyFormat.CslJson).Document;
        CslStyle style = CslStyle.Parse(Header + "<locale><terms><term name=\"open-quote\"></term><term name=\"close-quote\"></term></terms></locale><citation><layout><text variable=\"title\"/></layout></citation><bibliography><layout><text variable=\"title\"/></layout></bibliography></style>");
        var cluster = new CslCitation("cluster"); cluster.Items.Add(new CslCitationItem("a"));
        CslRenderResult result = new CslProcessor(data, style, new CslRenderOptions { OutputFormat = html ? CslOutputFormat.Html : CslOutputFormat.PlainText }).Render(new[] { cluster });
        Assert.Equal("", result.Citations.Single().Content);
        Assert.True(result.Citations.Single().IsEmpty);
        Assert.Equal(html ? "<div class=\"csl-entry\"></div>" : "", result.Bibliography.Single().Content);
        Assert.True(result.Bibliography.Single().IsEmpty);
        Assert.DoesNotContain(result.Bibliography, entry => !entry.IsEmpty);
    }

    [Fact]
    public void EmptyAndRemovedCitationClustersRemainDistinctInDocumentSnapshots() {
        BibliographyDocument data = BibliographyDocument.Parse("[{\"id\":\"a\",\"type\":\"book\",\"title\":\"Alpha\"}]", BibliographyFormat.CslJson).Document;
        CslStyle style = CslStyle.Parse(Header + "<citation><layout><text variable=\"title\"/></layout></citation></style>");
        var empty = new CslCitation("empty"); var present = new CslCitation("present"); present.Items.Add(new CslCitationItem("a"));
        var processor = new CslProcessor(data, style);
        CslRenderResult first = processor.Render(new[] { empty, present });
        Assert.Equal(new[] { "empty", "present" }, first.Citations.Select(entry => entry.Key));
        Assert.Equal(new[] { true, false }, first.Citations.Select(entry => entry.IsEmpty));
        CslRenderResult second = processor.Render(new[] { present });
        Assert.Equal("present", second.Citations.Single().Key);
        Assert.False(second.Citations.Single().IsEmpty);
    }
}
