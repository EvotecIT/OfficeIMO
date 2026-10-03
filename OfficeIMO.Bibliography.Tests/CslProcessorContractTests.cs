namespace OfficeIMO.Bibliography.Tests;

public sealed class CslProcessorContractTests {
    private const string Header = "<style xmlns=\"http://purl.org/net/xbiblio/csl\" version=\"1.0\" class=\"in-text\"><info><id>urn:test:style</id><title>Contract style</title><updated>2026-10-02T00:00:00Z</updated></info>";

    [Fact]
    public void CallerStyleMacrosLocaleRichTextAndSortingRenderThroughPublicBoundary() {
        BibliographyDocument document = BibliographyDocument.Parse("[{\"id\":\"b\",\"type\":\"book\",\"title\":\"<i>SCIENCE</i> &amp; <span class=\\\"nocase\\\">iPhone</span>\",\"author\":[{\"family\":\"Doe\",\"given\":\"Jane Mary\"}],\"issued\":{\"date-parts\":[[2025,5,15],[2025,5,17]]}},{\"id\":\"a\",\"type\":\"book\",\"title\":\"Alpha\",\"author\":[{\"family\":\"Adams\",\"given\":\"John\"}],\"issued\":{\"date-parts\":[[2024]]}}]", BibliographyFormat.CslJson).Document;
        CslStyle style = CslStyle.Parse(Header + "<macro name=\"author\"><names variable=\"author\"><name name-as-sort-order=\"all\" initialize-with=\". \"/></names></macro><citation><layout prefix=\"(\" suffix=\")\" delimiter=\"; \"><text macro=\"author\"/><date variable=\"issued\" prefix=\", \"><date-part name=\"year\"/></date></layout></citation><bibliography><sort><key macro=\"author\"/></sort><layout><group delimiter=\". \"><text macro=\"author\"/><text variable=\"title\" text-case=\"sentence\"/><date variable=\"issued\" form=\"text\"/></group></layout></bibliography></style>");
        var processor = new CslProcessor(document, style, new CslRenderOptions { OutputFormat = CslOutputFormat.Html });
        var cite = new CslCitation("cluster"); cite.Items.Add(new CslCitationItem("b"));
        CslRenderResult result = processor.Render(new[] { cite }, true);
        Assert.Equal("(Doe, J. M., 2025)", result.Citations.Single().Content);
        Assert.Equal(new[] { "a", "b" }, result.Bibliography.Select(entry => entry.Key));
        Assert.Contains("<i>Science</i> &amp; iPhone", result.Bibliography.Last().Content);
        Assert.Contains("May 15–17, 2025", result.Bibliography.Last().Content);
        Assert.False(document.IsModified);
    }

    [Fact]
    public void ReplacingDocumentCitationsRecalculatesAmbiguityAndNumericOrdering() {
        BibliographyDocument document = BibliographyDocument.Parse("[{\"id\":\"one\",\"type\":\"book\",\"title\":\"Alpha\",\"author\":[{\"family\":\"Doe\",\"given\":\"Jane\"}],\"issued\":{\"date-parts\":[[2025]]}},{\"id\":\"two\",\"type\":\"book\",\"title\":\"Beta\",\"author\":[{\"family\":\"Doe\",\"given\":\"Jane\"}],\"issued\":{\"date-parts\":[[2025]]}}]", BibliographyFormat.CslJson).Document;
        CslStyle style = CslStyle.Parse(Header + "<citation disambiguate-add-year-suffix=\"true\"><layout prefix=\"(\" suffix=\")\" delimiter=\"; \"><names variable=\"author\"><name form=\"short\"/></names><date variable=\"issued\" prefix=\" \"><date-part name=\"year\"/></date></layout></citation><bibliography><sort><key variable=\"title\"/></sort><layout><text variable=\"citation-number\" suffix=\". \"/><text variable=\"title\"/></layout></bibliography></style>");
        var processor = new CslProcessor(document, style);
        var cite = new CslCitation("cluster"); cite.Items.Add(new CslCitationItem("two")); cite.Items.Add(new CslCitationItem("one"));
        CslRenderResult result = processor.Render(new[] { cite });
        Assert.Equal("(Doe 2025b; Doe 2025a)", result.Citations.Single().Content);
        Assert.Equal(new[] { "1. Alpha", "2. Beta" }, result.Bibliography.Select(entry => entry.Content));
        cite.Items.RemoveAt(1);
        Assert.Equal("(Doe 2025)", processor.Render(new[] { cite }).Citations.Single().Content);
    }

    [Fact]
    public void NotePositionAndLocatorChangesAreRecomputedFromDocumentOrder() {
        BibliographyDocument document = BibliographyDocument.Parse("[{\"id\":\"one\",\"type\":\"book\",\"title\":\"Title\"}]", BibliographyFormat.CslJson).Document;
        CslStyle style = CslStyle.Parse(Header.Replace("in-text", "note") + "<citation><layout><choose><if position=\"ibid ibid-with-locator\" match=\"any\"><text term=\"ibid\"/></if><else><text variable=\"title\"/></else></choose><group prefix=\", \" delimiter=\" \"><label variable=\"locator\" form=\"short\"/><text variable=\"locator\"/></group></layout></citation></style>");
        CslCitation Cite(string id, int note, string locator) { var result = new CslCitation(id) { NoteIndex = note }; result.Items.Add(new CslCitationItem("one") { Locator = locator }); return result; }
        var citations = new[] { Cite("first", 1, "12"), Cite("repeat", 2, "15") };
        var processor = new CslProcessor(document, style);
        Assert.Equal(new[] { "Title, p. 12", "Ibid., p. 15" }, processor.Render(citations).Citations.Select(entry => entry.Content));
        Assert.Equal("Title, p. 15", processor.Render(citations.Skip(1)).Citations.Single().Content);
    }

    [Fact]
    public void EmptyVariableSuppressesGroupTermsAndAffixes() {
        BibliographyDocument document = BibliographyDocument.Parse("[{\"id\":\"one\",\"type\":\"book\",\"title\":\"Title\"}]", BibliographyFormat.CslJson).Document;
        CslStyle style = CslStyle.Parse(Header + "<citation><layout><text variable=\"title\"/><group prefix=\", \" delimiter=\" \"><text term=\"volume\"/><number variable=\"volume\"/></group></layout></citation></style>");
        var citation = new CslCitation("one"); citation.Items.Add(new CslCitationItem("one"));
        Assert.Equal("Title", new CslProcessor(document, style).Render(new[] { citation }).Citations.Single().Content);
    }

    [Fact]
    public void DependentStyleUsesLocalResolverAndLocaleOverride() {
        string parent = Header + "<citation><layout><text term=\"and\"/></layout></citation></style>";
        CslStyle style = CslStyle.Parse("<style xmlns=\"http://purl.org/net/xbiblio/csl\" version=\"1.0\" default-locale=\"pl-PL\"><info><link rel=\"independent-parent\" href=\"urn:test:parent\"/></info></style>", new CslStyleLoadOptions { IndependentStyleResolver = id => id == "urn:test:parent" ? parent : null });
        var options = new CslRenderOptions();
        options.Locales.Add("pl-PL", "<locale xmlns=\"http://purl.org/net/xbiblio/csl\" xml:lang=\"pl-PL\"><terms><term name=\"and\">i</term></terms></locale>");
        BibliographyDocument document = BibliographyDocument.Parse("[{\"id\":\"one\",\"type\":\"book\"}]", BibliographyFormat.CslJson).Document;
        var citation = new CslCitation("one"); citation.Items.Add(new CslCitationItem("one"));
        Assert.Equal("i", new CslProcessor(document, style, options).Render(new[] { citation }).Citations.Single().Content);
    }

    [Fact]
    public void ResourceLimitsCancellationAndInvalidReferencesFailBeforeReturningOutput() {
        Assert.Throws<System.Xml.XmlException>(() => CslStyle.Parse("<!DOCTYPE style [<!ENTITY external SYSTEM 'file:///missing'>]>" + Header + "<citation><layout><text value=\"&external;\"/></layout></citation></style>"));
        Assert.Throws<InvalidDataException>(() => CslStyle.Parse(Header + "<macro name=\"cycle\"><text macro=\"cycle\"/></macro><citation><layout><text macro=\"cycle\"/></layout></citation></style>"));
        Assert.Throws<InvalidDataException>(() => CslStyle.Parse(Header, new CslStyleLoadOptions { MaximumCharacters = 10 }));
        BibliographyDocument document = BibliographyDocument.Parse("[{\"id\":\"one\",\"type\":\"book\",\"title\":\"Long title\"}]", BibliographyFormat.CslJson).Document;
        CslStyle style = CslStyle.Parse(Header + "<citation><layout><text variable=\"title\"/></layout></citation></style>");
        var citation = new CslCitation("one"); citation.Items.Add(new CslCitationItem("one"));
        var processor = new CslProcessor(document, style, new CslRenderOptions { MaximumOutputCharacters = 5 });
        Assert.Throws<InvalidDataException>(() => processor.Render(new[] { citation }));
        Assert.Throws<OperationCanceledException>(() => processor.Render(new[] { citation }, cancellationToken: new CancellationToken(true)));
        citation.Items.Add(new CslCitationItem("missing"));
        Assert.Throws<ArgumentException>(() => processor.Render(new[] { citation }));
    }

    [Fact]
    public void RepeatedNonrecursiveMacroExpansionIsBoundedEvenWhenItProducesNoText() {
        var macros = new System.Text.StringBuilder("<macro name=\"m0\"><text variable=\"title\"/></macro>");
        for (int index = 1; index < 25; index++) macros.Append("<macro name=\"m").Append(index).Append("\"><text macro=\"m").Append(index - 1).Append("\"/><text macro=\"m").Append(index - 1).Append("\"/></macro>");
        CslStyle style = CslStyle.Parse(Header + macros + "<citation><layout><text macro=\"m24\"/></layout></citation></style>", new CslStyleLoadOptions { MaximumNestingDepth = 128 });
        BibliographyDocument document = BibliographyDocument.Parse("[{\"id\":\"one\",\"type\":\"book\"}]", BibliographyFormat.CslJson).Document;
        var cite = new CslCitation("cluster"); cite.Items.Add(new CslCitationItem("one"));
        var processor = new CslProcessor(document, style, new CslRenderOptions { MaximumRenderingOperations = 100 });
        InvalidDataException error = Assert.Throws<InvalidDataException>(() => processor.Render(new[] { cite }));
        Assert.Contains("MaximumRenderingOperations", error.Message);
    }
}
