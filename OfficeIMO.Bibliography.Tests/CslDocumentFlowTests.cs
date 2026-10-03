namespace OfficeIMO.Bibliography.Tests;

public sealed class CslDocumentFlowTests {
    private const string Header = "<style xmlns=\"http://purl.org/net/xbiblio/csl\" version=\"1.0\" class=\"in-text\">";
    private static BibliographyDocument Data(string json) => BibliographyDocument.Parse(json, BibliographyFormat.CslJson).Document;
    private static CslCitation Citation(params string[] keys) {
        var cite = new CslCitation("cluster");
        foreach (string key in keys) cite.Items.Add(new CslCitationItem(key));
        return cite;
    }

    [Fact]
    public void NumericRangesKeepEndpointAffixesAndFormatTheWholeCluster() {
        BibliographyDocument data = Data("[{\"id\":\"a\",\"type\":\"book\"},{\"id\":\"b\",\"type\":\"book\"},{\"id\":\"c\",\"type\":\"book\"},{\"id\":\"d\",\"type\":\"book\"}]");
        CslStyle style = CslStyle.Parse(Header + "<citation collapse=\"citation-number\"><layout font-weight=\"bold\" prefix=\"(\" suffix=\")\" delimiter=\", \"><text variable=\"citation-number\" prefix=\"[\" suffix=\"]\"/><text variable=\"locator\" prefix=\", \"/></layout></citation></style>");
        var processor = new CslProcessor(data, style, new CslRenderOptions { OutputFormat = CslOutputFormat.Html });
        CslCitation cite = Citation("a", "b", "c", "d");
        Assert.Equal("<b>([1]–[4])</b>", processor.Render(new[] { cite }).Citations.Single().Content);
        cite.Items[1].Locator = "12";
        Assert.Equal("<b>([1], [2], 12, [3], [4])</b>", processor.Render(new[] { cite }).Citations.Single().Content);
    }

    [Fact]
    public void GroupingAndYearSuffixRangesPreserveOtherAuthorGroups() {
        BibliographyDocument data = Data("[{\"id\":\"a\",\"type\":\"book\",\"title\":\"A\",\"author\":[{\"family\":\"Doe\"}],\"issued\":{\"date-parts\":[[2025]]}},{\"id\":\"b\",\"type\":\"book\",\"title\":\"B\",\"author\":[{\"family\":\"Doe\"}],\"issued\":{\"date-parts\":[[2025]]}},{\"id\":\"c\",\"type\":\"book\",\"title\":\"C\",\"author\":[{\"family\":\"Doe\"}],\"issued\":{\"date-parts\":[[2025]]}},{\"id\":\"d\",\"type\":\"book\",\"author\":[{\"family\":\"Adams\"}],\"issued\":{\"date-parts\":[[2024]]}}]");
        CslStyle style = CslStyle.Parse(Header + "<citation collapse=\"year-suffix-ranged\" disambiguate-add-year-suffix=\"true\" year-suffix-delimiter=\",\" cite-group-delimiter=\", \" after-collapse-delimiter=\"; \"><layout prefix=\"(\" suffix=\")\" delimiter=\"; \"><names variable=\"author\"><name form=\"short\"/></names><date variable=\"issued\" prefix=\" \"><date-part name=\"year\"/></date></layout></citation><bibliography><sort><key variable=\"title\"/></sort><layout><text variable=\"title\"/></layout></bibliography></style>");
        var processor = new CslProcessor(data, style);
        Assert.Equal("(Doe 2025a–c; Adams 2024)", processor.Render(new[] { Citation("a", "d", "b", "c") }).Citations.Single().Content);
    }

    [Theory]
    [InlineData("year-suffix", "Doe 2025a,b;2026")]
    [InlineData("year-suffix-ranged", "Doe 2025a,b;2026")]
    public void AYearSuffixGroupUsesTheLayoutDelimiterBeforeTheNextYear(string mode, string expected) {
        BibliographyDocument data = Data("[{\"id\":\"a\",\"type\":\"book\",\"author\":[{\"family\":\"Doe\"}],\"issued\":{\"date-parts\":[[2025]]}},{\"id\":\"b\",\"type\":\"book\",\"author\":[{\"family\":\"Doe\"}],\"issued\":{\"date-parts\":[[2025]]}},{\"id\":\"c\",\"type\":\"book\",\"author\":[{\"family\":\"Doe\"}],\"issued\":{\"date-parts\":[[2026]]}}]");
        CslStyle style = CslStyle.Parse(Header + "<citation collapse=\"" + mode + "\" disambiguate-add-year-suffix=\"true\" year-suffix-delimiter=\",\"><layout delimiter=\";\"><group delimiter=\" \"><names variable=\"author\"><name form=\"short\"/></names><date variable=\"issued\"><date-part name=\"year\"/></date></group></layout></citation></style>");
        Assert.Equal(expected, new CslProcessor(data, style).Render(new[] { Citation("a", "b", "c") }).Citations.Single().Content);
    }

    [Fact]
    public void AnExplicitAfterCollapseDelimiterSeparatesCollapsedGroupsAndLocators() {
        BibliographyDocument data = Data("[{\"id\":\"a\",\"type\":\"book\",\"author\":[{\"family\":\"Doe\"}],\"issued\":{\"date-parts\":[[2025]]}},{\"id\":\"b\",\"type\":\"book\",\"author\":[{\"family\":\"Doe\"}],\"issued\":{\"date-parts\":[[2026]]}},{\"id\":\"c\",\"type\":\"book\",\"author\":[{\"family\":\"Roe\"}],\"issued\":{\"date-parts\":[[2024]]}}]");
        CslStyle style = CslStyle.Parse(Header + "<citation collapse=\"year\" after-collapse-delimiter=\"; \"><layout delimiter=\", \"><group delimiter=\" \"><names variable=\"author\"><name form=\"short\"/></names><date variable=\"issued\"><date-part name=\"year\"/></date></group><text variable=\"locator\" prefix=\", \"/></layout></citation></style>");
        var processor = new CslProcessor(data, style);
        CslCitation cites = Citation("a", "b", "c"); cites.Items[0].Locator = "12"; cites.Items[1].Locator = "24";
        Assert.Equal("Doe 2025, 12; 2026, 24; Roe 2024", processor.Render(new[] { cites }).Citations.Single().Content);
        Assert.Equal("Doe 2025, Roe 2024", processor.Render(new[] { Citation("a", "c") }).Citations.Single().Content);
    }

    [Theory]
    [InlineData("complete-all", "—", "Doe & Smith")]
    [InlineData("complete-each", "— & —", "Doe & Smith")]
    [InlineData("partial-each", "— & —", "— & Smith")]
    [InlineData("partial-first", "— & Roe", "— & Smith")]
    public void BibliographySubstitutionUsesThePreviousRenderedNameList(string rule, string second, string third) {
        BibliographyDocument data = Data("[{\"id\":\"a\",\"type\":\"book\",\"author\":[{\"family\":\"Doe\"},{\"family\":\"Roe\"}]},{\"id\":\"b\",\"type\":\"book\",\"author\":[{\"family\":\"Doe\"},{\"family\":\"Roe\"}]},{\"id\":\"c\",\"type\":\"book\",\"author\":[{\"family\":\"Doe\"},{\"family\":\"Smith\"}]}]");
        CslStyle style = CslStyle.Parse(Header + "<citation><layout><text variable=\"title\"/></layout></citation><bibliography subsequent-author-substitute=\"—\" subsequent-author-substitute-rule=\"" + rule + "\"><layout><names variable=\"author\"><name form=\"short\" and=\"symbol\"/></names></layout></bibliography></style>");
        Assert.Equal(new[] { "Doe & Roe", second, third }, new CslProcessor(data, style).RenderBibliography().Select(entry => entry.Content));
    }

    [Fact]
    public void DatesSortAcrossTheEraBoundaryAndOrderRangesAfterSingleDates() {
        BibliographyDocument data = Data("[{\"id\":\"ad\",\"type\":\"book\",\"issued\":{\"date-parts\":[[50]]}},{\"id\":\"bc\",\"type\":\"book\",\"issued\":{\"date-parts\":[[-100]]}},{\"id\":\"range\",\"type\":\"book\",\"issued\":{\"date-parts\":[[50],[60]]}},{\"id\":\"later-bc\",\"type\":\"book\",\"issued\":{\"date-parts\":[[-50]]}}]");
        CslStyle style = CslStyle.Parse(Header + "<citation><layout><text variable=\"title\"/></layout></citation><bibliography><sort><key variable=\"issued\"/></sort><layout><text variable=\"id\"/></layout></bibliography></style>");
        Assert.Equal(new[] { "bc", "later-bc", "ad", "range" }, new CslProcessor(data, style).RenderBibliography().Select(entry => entry.Key));
    }

    [Fact]
    public void SortKeyNameOptionsOverrideTheDisplayedAbbreviation() {
        BibliographyDocument data = Data("[{\"id\":\"a\",\"type\":\"book\",\"author\":[{\"family\":\"Adams\"},{\"family\":\"Roe\"},{\"family\":\"Smith\"}]},{\"id\":\"b\",\"type\":\"book\",\"author\":[{\"family\":\"Adams\"},{\"family\":\"Doe\"},{\"family\":\"Smith\"}]}]");
        CslStyle style = CslStyle.Parse(Header + "<macro name=\"author\"><names variable=\"author\"><name et-al-min=\"3\" et-al-use-first=\"1\"/></names></macro><citation><layout><text macro=\"author\"/></layout></citation><bibliography><sort><key macro=\"author\" names-min=\"3\" names-use-first=\"2\"/></sort><layout><text macro=\"author\"/></layout></bibliography></style>");
        Assert.Equal(new[] { "b", "a" }, new CslProcessor(data, style).RenderBibliography().Select(entry => entry.Key));
        Assert.All(new CslProcessor(data, style).RenderBibliography(), entry => Assert.Equal("Adams et al.", entry.Content));
    }

    [Theory]
    [InlineData(null, "505-517, 1496-1504, 321-328")]
    [InlineData("expanded", "505–517, 1496–1504, 321–328")]
    [InlineData("minimal", "505–17, 1496–504, 321–8")]
    [InlineData("minimal-two", "505–17, 1496–504, 321–28")]
    [InlineData("chicago", "505–17, 1496–1504, 321–28")]
    public void PageRangePolicyDoesNotAbbreviateCitationLocators(string? format, string expected) {
        BibliographyDocument data = Data("[{\"id\":\"a\",\"type\":\"book\",\"page\":\"505-517, 1496-1504, 321-328\"}]");
        CslStyle style = CslStyle.Parse(Header.Replace("class=\"in-text\"", "class=\"in-text\"" + (format == null ? string.Empty : " page-range-format=\"" + format + "\"")) + "<citation><layout><group delimiter=\"; \"><text variable=\"page\"/><text variable=\"locator\"/></group></layout></citation></style>");
        CslCitation cite = Citation("a"); cite.Items[0].Locator = "505-517";
        Assert.Equal(expected + "; 505–517", new CslProcessor(data, style).Render(new[] { cite }).Citations.Single().Content);
    }

    [Fact]
    public void IbidAppliesWithinClustersAndToTheFirstCiteAfterASingleCite() {
        BibliographyDocument data = Data("[{\"id\":\"a\",\"type\":\"book\"},{\"id\":\"b\",\"type\":\"book\"}]");
        CslStyle style = CslStyle.Parse(Header.Replace("in-text", "note") + "<citation><layout delimiter=\"; \"><choose><if position=\"ibid-with-locator\"><text value=\"changed\"/></if><else-if position=\"ibid\"><text value=\"ibid\"/></else-if><else><text value=\"full\"/></else></choose><text variable=\"first-reference-note-number\" prefix=\"@\"/></layout></citation></style>");
        CslCitation first = Citation("a"); first.NoteIndex = 1;
        CslCitation second = new CslCitation("second") { NoteIndex = 2 }; second.Items.Add(new CslCitationItem("a")); second.Items.Add(new CslCitationItem("a") { Locator = "15" }); second.Items.Add(new CslCitationItem("b"));
        Assert.Equal(new[] { "Full", "Ibid@1; changed@1; full" }, new CslProcessor(data, style).Render(new[] { first, second }).Citations.Select(entry => entry.Content));
    }
}
