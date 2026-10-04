namespace OfficeIMO.Bibliography.Tests;

public sealed class CslConditionalDetailTests {
    private const string Header = "<style xmlns=\"http://purl.org/net/xbiblio/csl\" version=\"1.0\" class=\"note\"><info><title>Conditional detail</title><id>urn:conditional-detail</id><updated>2026-10-03T00:00:00Z</updated></info>";
    private const string Author = "<names variable=\"author\"><name form=\"short\"/></names>";
    private const string Title = "<choose><if disambiguate=\"true\"><text variable=\"title\"/></if></choose>";
    private const string Edition = "<choose><if disambiguate=\"true\"><text variable=\"edition\" prefix=\"ed. \"/></if></choose>";
    private static BibliographyDocument Data(string first = "Alpha", string second = "Beta", bool sibling = false) => BibliographyDocument.Parse(
        "[{\"id\":\"a\",\"type\":\"book\",\"author\":[{\"family\":\"Doe\",\"given\":\"John\"}],\"title\":\"" + first + "\",\"edition\":\"1\",\"abstract\":\"Common\"},{\"id\":\"b\",\"type\":\"book\",\"author\":[{\"family\":\"Doe\",\"given\":\"John\"}],\"title\":\"" + second + "\",\"edition\":\"2\",\"abstract\":\"Common\"}" +
        (sibling ? ",{\"id\":\"c\",\"type\":\"book\",\"author\":[{\"family\":\"Doe\",\"given\":\"John\"}],\"title\":\"Gamma\",\"edition\":\"3\"}" : "") + "]", BibliographyFormat.CslJson).Document;

    private static CslCitation Cite(string id, int note, params string[] keys) {
        var citation = new CslCitation(id) { NoteIndex = note };
        foreach (string key in keys) citation.Items.Add(new CslCitationItem(key));
        return citation;
    }

    private static CslStyle Style(string fields, string macros = "", string options = "") => CslStyle.Parse(Header + macros +
        "<citation " + options + "><layout delimiter=\"; \"><group delimiter=\", \">" + fields + "</group></layout></citation></style>");

    [Fact]
    public void DistinguishingTitlesDoNotAlsoAddEditions() => Assert.Equal("Doe, Alpha; Doe, Beta",
        new CslProcessor(Data(), Style(Author + Title + Edition)).Render(new[] { Cite("one", 1, "a", "b") }).Citations.Single().Content);

    [Fact]
    public void EditionsAreAddedOnlyToTheRemainingAmbiguousWorks() => Assert.Equal("Doe, Shared, ed. 1; Doe, Shared, ed. 2; Doe, Gamma",
        new CslProcessor(Data("Shared", "Shared", true), Style(Author + Title + Edition)).Render(new[] { Cite("one", 1, "a", "b", "c") }).Citations.Single().Content);

    [Fact]
    public void UnhelpfulEarlierDetailIsOmittedWhenAnotherConditionDistinguishesTheWorks() => Assert.Equal("Doe, ed. 1; Doe, ed. 2",
        new CslProcessor(Data("Shared", "Shared"), Style(Author + Title + Edition)).Render(new[] { Cite("one", 1, "a", "b") }).Citations.Single().Content);

    [Fact]
    public void ReusedMacroCallsSelectOnlyTheNeededOccurrence() {
        string macros = "<macro name=\"detail\">" + Title + "</macro>";
        Assert.Equal("Doe, Alpha; Doe, Beta", new CslProcessor(Data(), Style(Author + "<text macro=\"detail\"/><text macro=\"detail\" prefix=\"duplicate=\"/>", macros))
            .Render(new[] { Cite("one", 1, "a", "b") }).Citations.Single().Content);
    }

    [Fact]
    public void NestedConditionsCanOpenTheGateNeededForDistinguishingDetail() {
        string nested = "<choose><if disambiguate=\"true\"><group delimiter=\" \"><text value=\"Details\"/>" + Title + Edition + "</group></if></choose>";
        Assert.Equal("Doe, Details Alpha; Doe, Details Beta", new CslProcessor(Data(), Style(Author + nested))
            .Render(new[] { Cite("one", 1, "a", "b") }).Citations.Single().Content);
    }

    [Fact]
    public void SelectedConditionalBranchReplacesItsAlternative() {
        string choice = "<choose><if disambiguate=\"true\"><text variable=\"title\"/></if><else><text variable=\"abstract\"/></else></choose>";
        Assert.Equal("Doe, Alpha; Doe, Beta", new CslProcessor(Data(), Style(Author + choice + Edition))
            .Render(new[] { Cite("one", 1, "a", "b") }).Citations.Single().Content);
    }

    [Fact]
    public void UniqueFullNotesDoNotReceiveDetailNeededOnlyByShortNotes() {
        string fields = "<choose><if position=\"first\"><names variable=\"author\"><name/></names><text variable=\"title\"/></if><else>" + Author + "</else></choose>" + Title;
        CslRenderResult result = new CslProcessor(Data(), Style(fields)).Render(new[] { Cite("first", 1, "a", "b"), Cite("short", 2, "a", "b") });
        Assert.Equal("John Doe, Alpha; John Doe, Beta", result.Citations[0].Content);
        Assert.Equal("Doe, Alpha; Doe, Beta", result.Citations[1].Content);
    }

    [Fact]
    public void ConditionalSelectionIsRecalculatedForEachDocumentSnapshot() {
        var processor = new CslProcessor(Data(), Style(Author + Title + Edition));
        Assert.Equal("Doe, Alpha; Doe, Beta", processor.Render(new[] { Cite("both", 1, "a", "b") }).Citations.Single().Content);
        Assert.Equal("Doe", processor.Render(new[] { Cite("one", 1, "a") }).Citations.Single().Content);
        Assert.Equal("Doe, Alpha; Doe, Beta", processor.Render(new[] { Cite("both", 1, "a", "b") }).Citations.Single().Content);
    }

    [Fact]
    public void ConditionsInCopiedNameSubstitutesRetainTheirCallIdentity() {
        string names = "<names variable=\"editor\"><substitute><names variable=\"translator\"><substitute><group delimiter=\", \"><text value=\"Doe\"/>" + Title + Edition + "</group></substitute></names></substitute></names>";
        Assert.Equal("Doe, Alpha; Doe, Beta", new CslProcessor(Data(), Style(names))
            .Render(new[] { Cite("one", 1, "a", "b") }).Citations.Single().Content);
    }

    [Fact]
    public void NewlyExposedCollisionsAlsoReceiveTheDistinguishingBranch() {
        var data = BibliographyDocument.Parse("[{\"id\":\"a\",\"type\":\"book\",\"abstract\":\"Common\",\"title\":\"Alpha\"},{\"id\":\"b\",\"type\":\"book\",\"abstract\":\"Common\",\"title\":\"Beta\"},{\"id\":\"c\",\"type\":\"book\",\"abstract\":\"Alpha\",\"title\":\"Gamma\"}]", BibliographyFormat.CslJson).Document;
        string choice = "<choose><if disambiguate=\"true\"><text variable=\"title\"/></if><else><text variable=\"abstract\"/></else></choose>";
        Assert.Equal("Alpha; Beta; Gamma", new CslProcessor(data, Style(choice)).Render(new[] { Cite("one", 1, "a", "b", "c") }).Citations.Single().Content);
    }

    [Fact]
    public void UnhelpfulNestedTrialsRollBackAllTheirChoices() {
        var data = BibliographyDocument.Parse("[{\"id\":\"a\",\"type\":\"book\",\"abstract\":\"Common\",\"title\":\"Shared\"},{\"id\":\"b\",\"type\":\"book\",\"abstract\":\"Common\",\"title\":\"Shared\"},{\"id\":\"c\",\"type\":\"book\",\"abstract\":\"Details\",\"title\":\"Unique\"}]", BibliographyFormat.CslJson).Document;
        string choice = "<choose><if disambiguate=\"true\"><group delimiter=\" \"><text value=\"Details\"/>" + Title + "</group></if><else><text variable=\"abstract\"/></else></choose>";
        Assert.Equal("Common; Common; Details", new CslProcessor(data, Style(choice)).Render(new[] { Cite("one", 1, "a", "b", "c") }).Citations.Single().Content);
    }

    [Theory]
    [InlineData("note")]
    [InlineData("in-text")]
    public void RepeatedCitationsReuseDetailWhenTheStyleHasOneReusableForm(string styleClass) {
        CslStyle style = CslStyle.Parse(Header.Replace("class=\"note\"", "class=\"" + styleClass + "\"") +
            "<citation><layout delimiter=\"; \"><group delimiter=\", \">" + Author + Title + Edition + "</group></layout></citation></style>");
        Assert.Equal(new[] { "Doe, Alpha; Doe, Beta", "Doe, Alpha; Doe, Beta" }, new CslProcessor(Data(), style)
            .Render(new[] { Cite("first", 1, "a", "b"), Cite("again", 2, "a", "b") }).Citations.Select(entry => entry.Content));
    }

    [Fact]
    public void ExplicitConditionalYearSuffixIsSelectedAfterSuffixAssignment() {
        var data = BibliographyDocument.Parse("[{\"id\":\"a\",\"type\":\"book\",\"author\":[{\"family\":\"Doe\"}],\"issued\":{\"date-parts\":[[2020]]}},{\"id\":\"b\",\"type\":\"book\",\"author\":[{\"family\":\"Doe\"}],\"issued\":{\"date-parts\":[[2020]]}}]", BibliographyFormat.CslJson).Document;
        string year = "<group><date variable=\"issued\"><date-part name=\"year\"/></date><choose><if disambiguate=\"true\"><text variable=\"year-suffix\"/></if></choose></group>";
        var processor = new CslProcessor(data, Style(Author + year, options: "disambiguate-add-year-suffix=\"true\""));
        Assert.Equal("Doe, 2020a; Doe, 2020b", processor.Render(new[] { Cite("one", 1, "a", "b") }).Citations.Single().Content);
        Assert.Equal("Doe, 2020", processor.Render(new[] { Cite("one", 1, "a") }).Citations.Single().Content);
    }

    [Theory]
    [InlineData("ascending", "Doe, 2020a; Doe, 2020b; Doe, 2020c; Doe, 2020d", "Xa,Xb,Yc,Yd")]
    [InlineData("descending", "Doe, 2020c; Doe, 2020d; Doe, 2020a; Doe, 2020b", "Ya,Yb,Xc,Xd")]
    public void ConditionalReplacementReconcilesSuffixesAcrossFormerlySeparateGroups(string direction, string expected, string bibliography) {
        string[] records = Enumerable.Range(0, 4).Select(index => "{\"id\":\"" + (char)('a' + index) + "\",\"type\":\"book\",\"title\":\"" + (index < 2 ? "X" : "Y") + "\",\"author\":[{\"family\":\"Doe\"}],\"issued\":{\"date-parts\":[[2020]]}}").ToArray();
        var data = BibliographyDocument.Parse("[" + string.Join(",", records) + "]", BibliographyFormat.CslJson).Document;
        string choice = "<choose><if disambiguate=\"true\"><group><date variable=\"issued\"><date-part name=\"year\"/></date><text variable=\"year-suffix\"/></group></if><else><text variable=\"title\"/></else></choose>";
        var style = CslStyle.Parse(Header + "<citation disambiguate-add-year-suffix=\"true\"><layout delimiter=\"; \"><group delimiter=\", \">" + Author + choice + "</group></layout></citation><bibliography><sort><key variable=\"title\" sort=\"" + direction + "\"/></sort><layout><text variable=\"title\"/><text variable=\"year-suffix\"/></layout></bibliography></style>");
        var processor = new CslProcessor(data, style);
        CslRenderResult result = processor.Render(new[] { Cite("one", 1, "a", "b", "c", "d") });
        Assert.Equal(expected, result.Citations.Single().Content);
        Assert.Equal(bibliography.Split(','), result.Bibliography.Select(entry => entry.Content));
        Assert.Equal("Doe, X; Doe, Y", processor.Render(new[] { Cite("one", 1, "a", "c") }).Citations.Single().Content);
    }

    [Fact]
    public void SiblingConditionsCanSupplyDetailThatOnlyTheirCombinationRenders() {
        var data = BibliographyDocument.Parse("[{\"id\":\"a\",\"type\":\"book\",\"title\":\"Shared\"},{\"id\":\"b\",\"type\":\"article-journal\",\"title\":\"Shared\"}]", BibliographyFormat.CslJson).Document;
        string kind = "<choose><if disambiguate=\"true\"><choose><if type=\"book\"><text value=\"Book\"/></if><else><text value=\"Article\"/></else></choose></if></choose>";
        string redundant = "<choose><if disambiguate=\"true\"><text value=\"Unneeded\"/></if></choose>";
        Assert.Equal("Book, Shared; Article, Shared", new CslProcessor(data, Style(redundant + kind + Title + "<text variable=\"abstract\"/>"))
            .Render(new[] { Cite("one", 1, "a", "b") }).Citations.Single().Content);
    }

    [Fact]
    public void ConditionalAttemptsConsumeTheSharedRenderingWorkBudget() {
        string constants = string.Concat(Enumerable.Range(0, 100).Select(index => "<choose><if disambiguate=\"true\"><text value=\"same\"/></if></choose>"));
        var processor = new CslProcessor(Data(), Style(Author + constants + Edition), new CslRenderOptions { MaximumRenderingOperations = 10000 });
        Assert.Throws<InvalidDataException>(() => processor.Render(new[] { Cite("one", 1, "a", "b") }));
    }
}
