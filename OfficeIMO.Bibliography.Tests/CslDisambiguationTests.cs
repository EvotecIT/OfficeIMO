namespace OfficeIMO.Bibliography.Tests;

public sealed class CslDisambiguationTests {
    private const string Header = "<style xmlns=\"http://purl.org/net/xbiblio/csl\" version=\"1.0\" class=\"in-text\">";

    [Theory]
    [InlineData("[[2020],[2021]]", "<date-part name=\"year\"/>", "2020a–2021; 2020b–2021")]
    [InlineData("[[2020,1,1],[2020,1,2]]", "<date-part name=\"month\" form=\"short\" suffix=\" \"/><date-part name=\"day\" suffix=\", \"/><date-part name=\"year\"/>", "Jan. 1–2, 2020a; Jan. 1–2, 2020b")]
    [InlineData("[[2020,1,1],[2020,1,2]]", "<date-part name=\"year\" suffix=\" \"/><date-part name=\"month\" form=\"short\" suffix=\" \"/><date-part name=\"day\"/>", "2020a Jan. 1–2; 2020b Jan. 1–2")]
    public void AutomaticYearSuffixAppearsOnTheFirstRenderedYearWithoutCreatingAnArtificialRange(string parts, string dateStyle, string expected) {
        BibliographyDocument data = BibliographyDocument.Parse("[{\"id\":\"a\",\"type\":\"book\",\"issued\":{\"date-parts\":" + parts + "}}," +
            "{\"id\":\"b\",\"type\":\"book\",\"issued\":{\"date-parts\":" + parts + "}}]", BibliographyFormat.CslJson).Document;
        CslStyle style = CslStyle.Parse(Header + "<citation disambiguate-add-year-suffix=\"true\"><layout delimiter=\"; \"><date variable=\"issued\">" + dateStyle + "</date></layout></citation></style>");
        var citation = new CslCitation("cluster");
        citation.Items.Add(new CslCitationItem("a"));
        citation.Items.Add(new CslCitationItem("b"));
        Assert.Equal(expected, new CslProcessor(data, style).Render(new[] { citation }).Citations.Single().Content);
    }
    private static BibliographyDocument Data(string firstGiven, string secondGiven, int firstYear, int secondYear) => BibliographyDocument.Parse(
        "[{\"id\":\"a\",\"type\":\"book\",\"author\":[{\"family\":\"Doe\",\"given\":\"" + firstGiven + "\"}],\"issued\":{\"date-parts\":[[" + firstYear + "]]}},{\"id\":\"b\",\"type\":\"book\",\"author\":[{\"family\":\"Doe\",\"given\":\"" + secondGiven + "\"}],\"issued\":{\"date-parts\":[[" + secondYear + "]]}}]", BibliographyFormat.CslJson).Document;

    private static string Render(BibliographyDocument data, string rule) {
        CslStyle style = CslStyle.Parse(Header + "<citation disambiguate-add-givenname=\"true\" givenname-disambiguation-rule=\"" + rule + "\"><layout delimiter=\"; \"><names variable=\"author\"><name form=\"short\" initialize-with=\". \"/></names><date variable=\"issued\" prefix=\" \"><date-part name=\"year\"/></date></layout></citation></style>");
        var cite = new CslCitation("cluster"); cite.Items.Add(new CslCitationItem("a")); cite.Items.Add(new CslCitationItem("b"));
        return new CslProcessor(data, style).Render(new[] { cite }).Citations.Single().Content;
    }

    [Theory]
    [InlineData("by-cite", "Doe 2020; Doe 2021")]
    [InlineData("all-names", "J. Doe 2020; R. Doe 2021")]
    [InlineData("all-names-with-initials", "J. Doe 2020; R. Doe 2021")]
    [InlineData("primary-name", "J. Doe 2020; R. Doe 2021")]
    [InlineData("primary-name-with-initials", "J. Doe 2020; R. Doe 2021")]
    public void GlobalNameRulesExpandEvenWhenYearsAlreadyDistinguishCites(string rule, string expected) =>
        Assert.Equal(expected, Render(Data("John", "Ruth", 2020, 2021), rule));

    [Theory]
    [InlineData("by-cite", "John Doe 2025; Jane Doe 2025")]
    [InlineData("all-names-with-initials", "Doe 2025; Doe 2025")]
    public void UnsuccessfulInitialExpansionKeepsTheOriginalNamesUnlessFullNamesAreAllowed(string rule, string expected) =>
        Assert.Equal(expected, Render(Data("John", "Jane", 2025, 2025), rule));

    [Fact]
    public void ShowingMoreIdenticalAuthorsDoesNotLeaveAnUnsuccessfulExpansionInOutput() {
        BibliographyDocument document = BibliographyDocument.Parse("[{\"id\":\"a\",\"type\":\"book\",\"author\":[{\"family\":\"Doe\"},{\"family\":\"Roe\"},{\"family\":\"Smith\"}]},{\"id\":\"b\",\"type\":\"book\",\"author\":[{\"family\":\"Doe\"},{\"family\":\"Roe\"},{\"family\":\"Smith\"}]}]", BibliographyFormat.CslJson).Document;
        CslStyle style = CslStyle.Parse(Header + "<citation disambiguate-add-names=\"true\"><layout delimiter=\"; \"><names variable=\"author\"><name form=\"short\" et-al-min=\"3\" et-al-use-first=\"1\"/></names></layout></citation></style>");
        var cite = new CslCitation("cluster"); cite.Items.Add(new CslCitationItem("a")); cite.Items.Add(new CslCitationItem("b"));
        Assert.Equal("Doe et al.; Doe et al.", new CslProcessor(document, style).Render(new[] { cite }).Citations.Single().Content);
    }

    [Fact]
    public void IdenticalReferencesKeepEnoughAuthorsToDistinguishADifferentSiblingReference() {
        BibliographyDocument document = BibliographyDocument.Parse("[{\"id\":\"a\",\"type\":\"book\",\"author\":[{\"family\":\"Smith\"},{\"family\":\"Brown\"},{\"family\":\"Jones\"}]},{\"id\":\"b\",\"type\":\"book\",\"author\":[{\"family\":\"Smith\"},{\"family\":\"Brown\"},{\"family\":\"Jones\"}]},{\"id\":\"c\",\"type\":\"book\",\"author\":[{\"family\":\"Smith\"},{\"family\":\"Benson\"},{\"family\":\"Jones\"}]}]", BibliographyFormat.CslJson).Document;
        CslStyle style = CslStyle.Parse(Header + "<citation disambiguate-add-names=\"true\"><layout delimiter=\"; \"><names variable=\"author\"><name form=\"short\" et-al-min=\"3\" et-al-use-first=\"1\"/></names></layout></citation></style>");
        var cite = new CslCitation("cluster"); foreach (string key in new[] { "a", "b", "c" }) cite.Items.Add(new CslCitationItem(key));
        Assert.Equal("Smith, Brown, et al.; Smith, Brown, et al.; Smith, Benson, et al.", new CslProcessor(document, style).Render(new[] { cite }).Citations.Single().Content);
    }
}
