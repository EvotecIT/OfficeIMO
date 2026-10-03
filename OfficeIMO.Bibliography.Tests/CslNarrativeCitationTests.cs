namespace OfficeIMO.Bibliography.Tests;

public sealed class CslNarrativeCitationTests {
    private const string Header = "<style xmlns=\"http://purl.org/net/xbiblio/csl\" version=\"1.0\" class=\"in-text\">";

    [Theory]
    [InlineData("\"author\":[{\"family\":\"Doe\",\"given\":\"Jane\"}]", "<names variable=\"author\"><name form=\"short\" font-style=\"italic\"/></names>", "<i>Doe</i>")]
    [InlineData("\"editor\":[{\"family\":\"Editor\",\"given\":\"Emily\"}]", "<names variable=\"author\" font-weight=\"bold\"><name form=\"short\"/><substitute><names variable=\"editor\"/></substitute></names>", "<b>Editor</b>")]
    [InlineData("\"title\":\"Anonymous\"", "<names variable=\"author\"><name/><substitute><text variable=\"title\" font-style=\"italic\"/></substitute></names>", "<i>Anonymous</i>")]
    public void NarrativeCitationsUseStyleNamesAndSubstitutionsWithoutLayoutParentheses(string fields, string names, string expected) {
        var document = BibliographyDocument.Parse("[{\"id\":\"a\",\"type\":\"book\",\"issued\":{\"date-parts\":[[2026]]}," + fields + "}]", BibliographyFormat.CslJson).Document;
        CslStyle style = CslStyle.Parse(Header + "<macro name=\"creator\">" + names + "</macro><citation><layout prefix=\"(\" suffix=\")\"><text macro=\"creator\"/><date variable=\"issued\" prefix=\", \"><date-part name=\"year\"/></date></layout></citation></style>");
        var cite = new CslCitation("cluster"); cite.Items.Add(new CslCitationItem("a") { AuthorOnly = true });
        Assert.Equal(expected, new CslProcessor(document, style, new CslRenderOptions { OutputFormat = CslOutputFormat.Html }).Render(new[] { cite }).Citations.Single().Content);
    }

    [Fact]
    public void NumericStylesUseAuthorFallbackAndContradictoryFlagsAreRejected() {
        var document = BibliographyDocument.Parse("[{\"id\":\"a\",\"type\":\"book\",\"author\":[{\"family\":\"Doe\",\"given\":\"Jane\"}]}]", BibliographyFormat.CslJson).Document;
        CslStyle style = CslStyle.Parse(Header + "<citation><layout prefix=\"[\" suffix=\"]\"><text variable=\"citation-number\"/></layout></citation></style>");
        var cite = new CslCitation("cluster"); cite.Items.Add(new CslCitationItem("a") { AuthorOnly = true });
        var processor = new CslProcessor(document, style);
        Assert.Equal("Jane Doe", processor.Render(new[] { cite }).Citations.Single().Content);
        cite.Items[0].SuppressAuthor = true;
        Assert.Throws<ArgumentException>(() => processor.Render(new[] { cite }));
    }

    [Fact]
    public void SortWorkSharesTheRenderingBudgetAndPreservesThePublicLimitException() {
        string data = "[" + string.Join(",", Enumerable.Range(0, 100).Select(index => "{\"id\":\"" + index + "\",\"type\":\"book\",\"title\":\"Title " + (100 - index) + "\"}")) + "]";
        CslStyle style = CslStyle.Parse(Header + "<citation><layout><text variable=\"title\"/></layout></citation><bibliography><sort><key variable=\"title\"/></sort><layout><text variable=\"title\"/></layout></bibliography></style>");
        var processor = new CslProcessor(BibliographyDocument.Parse(data, BibliographyFormat.CslJson).Document, style, new CslRenderOptions { MaximumRenderingOperations = 20 });
        InvalidDataException failure = Assert.Throws<InvalidDataException>(() => processor.Render(Array.Empty<CslCitation>(), includeUncitedItems: true));
        Assert.Contains("MaximumRenderingOperations", failure.Message);
    }
}
