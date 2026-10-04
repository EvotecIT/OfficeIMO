namespace OfficeIMO.Bibliography.Tests;

public sealed class CslNoteDisambiguationTests {
    private const string Header = "<style xmlns=\"http://purl.org/net/xbiblio/csl\" version=\"1.0\" class=\"note\"><info><title>Note contracts</title><id>urn:note-contracts</id><updated>2026-10-03T00:00:00Z</updated></info>";
    private const string First = "<names variable=\"author\"><name/></names><text variable=\"title\" prefix=\", \"/>";
    private const string Short = "<names variable=\"author\"><name form=\"short\"/></names><choose><if disambiguate=\"true\"><text variable=\"title\" prefix=\", \"/></if></choose>";

    private static BibliographyDocument Data(string firstNames = "[{\"family\":\"Doe\",\"given\":\"John\"}]", string secondNames = "[{\"family\":\"Doe\",\"given\":\"John\"}]") =>
        BibliographyDocument.Parse("[{\"id\":\"a\",\"type\":\"book\",\"title\":\"Alpha\",\"issued\":{\"date-parts\":[[2020]]},\"author\":" + firstNames + "},{\"id\":\"b\",\"type\":\"book\",\"title\":\"Beta\",\"issued\":{\"date-parts\":[[2020]]},\"author\":" + secondNames + "}]", BibliographyFormat.CslJson).Document;

    private static CslCitation Cite(string id, int note, params string[] keys) {
        var citation = new CslCitation(id) { NoteIndex = note };
        foreach (string key in keys) citation.Items.Add(new CslCitationItem(key));
        return citation;
    }

    private static CslStyle Style(string shortForm, string options = "", string bibliography = "") => CslStyle.Parse(Header +
        "<citation " + options + "><layout delimiter=\"; \"><choose><if position=\"first\">" + First + "</if><else>" + shortForm + "</else></choose></layout></citation>" + bibliography + "</style>");

    [Fact]
    public void DistinctFullNotesStillDisambiguateTheirShortForms() {
        var processor = new CslProcessor(Data(), Style(Short));
        CslRenderResult result = processor.Render(new[] { Cite("first", 1, "a", "b"), Cite("short", 2, "a", "b") });
        Assert.Equal("John Doe, Alpha; John Doe, Beta", result.Citations[0].Content);
        Assert.Equal("Doe, Alpha; Doe, Beta", result.Citations[1].Content);
    }

    [Theory]
    [InlineData(false, "Doe, supra 1; Doe, supra 2")]
    [InlineData(true, "Doe, Alpha, supra 1; Doe, Beta, supra 1")]
    public void NoteBackreferencesParticipateInTheAmbiguityComparison(bool sameFirstNote, string expected) {
        string shortForm = Short + "<text value=\", supra \"/><text variable=\"first-reference-note-number\"/>";
        CslCitation[] citations = sameFirstNote ? new[] { Cite("first", 1, "a", "b"), Cite("short", 2, "a", "b") } :
            new[] { Cite("first-a", 1, "a"), Cite("first-b", 2, "b"), Cite("short", 3, "a", "b") };
        Assert.Equal(expected, new CslProcessor(Data(), Style(shortForm)).Render(citations).Citations.Last().Content);
    }

    [Fact]
    public void NearNoteFormReceivesTheRequiredTitle() {
        string shortForm = "<choose><if position=\"near-note\">" + Short + "</if><else><text variable=\"title\"/></else></choose>";
        Assert.Equal("Doe, Alpha; Doe, Beta", new CslProcessor(Data(), Style(shortForm)).Render(new[] {
            Cite("first", 1, "a", "b"), Cite("near", 2, "a", "b") }).Citations.Last().Content);
    }

    [Fact]
    public void GivenNameExpansionAlsoExaminesTheShortNoteForm() {
        string shortForm = "<names variable=\"author\"><name form=\"short\" initialize-with=\". \"/></names>";
        Assert.Equal("J. Doe; R. Doe", new CslProcessor(Data(secondNames: "[{\"family\":\"Doe\",\"given\":\"Ruth\"}]"),
            Style(shortForm, "disambiguate-add-givenname=\"true\"")).Render(new[] {
            Cite("first", 1, "a", "b"), Cite("short", 2, "a", "b") }).Citations.Last().Content);
    }

    [Fact]
    public void EtAlExpansionAlsoExaminesTheShortNoteForm() {
        string shortForm = "<names variable=\"author\"><name form=\"short\" et-al-min=\"2\" et-al-use-first=\"1\"/></names>";
        BibliographyDocument data = Data("[{\"family\":\"Doe\"},{\"family\":\"Roe\"}]", "[{\"family\":\"Doe\"},{\"family\":\"Smith\"}]");
        Assert.Equal("Doe, Roe; Doe, Smith", new CslProcessor(data, Style(shortForm, "disambiguate-add-names=\"true\"")).Render(new[] {
            Cite("first", 1, "a", "b"), Cite("short", 2, "a", "b") }).Citations.Last().Content);
    }

    [Fact]
    public void ShortNoteYearSuffixesRemainConsistentWithTheBibliography() {
        string shortForm = "<names variable=\"author\"><name form=\"short\"/></names><date variable=\"issued\" prefix=\" \"><date-part name=\"year\"/></date>";
        string bibliography = "<bibliography><layout><text variable=\"title\" suffix=\" \"/><date variable=\"issued\"><date-part name=\"year\"/></date></layout></bibliography>";
        CslRenderResult result = new CslProcessor(Data(), Style(shortForm, "disambiguate-add-year-suffix=\"true\"", bibliography)).Render(new[] {
            Cite("first", 1, "a", "b"), Cite("short", 2, "a", "b") });
        Assert.Equal("Doe 2020a; Doe 2020b", result.Citations.Last().Content);
        Assert.Equal(new[] { "Alpha 2020a", "Beta 2020b" }, result.Bibliography.Select(entry => entry.Content));
    }

    [Fact]
    public void RemovingTheOtherBookRecalculatesShortNoteDisambiguation() {
        var processor = new CslProcessor(Data(), Style(Short));
        Assert.Equal("Doe, Alpha; Doe, Beta", processor.Render(new[] { Cite("first", 1, "a", "b"), Cite("short", 2, "a", "b") }).Citations.Last().Content);
        Assert.Equal("Doe", processor.Render(new[] { Cite("first", 1, "a"), Cite("short", 2, "a") }).Citations.Last().Content);
    }

    [Fact]
    public void OverlappingAmbiguitySetsShareOneBibliographyOrderedSuffixAssignment() {
        BibliographyDocument data = BibliographyDocument.Parse("[{\"id\":\"a\",\"type\":\"book\",\"title\":\"Shared\",\"author\":[{\"family\":\"Doe\"}]},{\"id\":\"b\",\"type\":\"book\",\"title\":\"Shared\",\"author\":[{\"family\":\"Roe\"}]},{\"id\":\"c\",\"type\":\"book\",\"title\":\"Unique\",\"author\":[{\"family\":\"Roe\"}]}]", BibliographyFormat.CslJson).Document;
        CslStyle style = CslStyle.Parse(Header + "<citation disambiguate-add-year-suffix=\"true\"><layout delimiter=\"; \"><choose><if position=\"first\"><text variable=\"title\"/></if><else><names variable=\"author\"><name form=\"short\"/></names></else></choose><text variable=\"year-suffix\"/></layout></citation></style>");
        CslRenderResult result = new CslProcessor(data, style).Render(new[] { Cite("first", 1, "a", "b", "c"), Cite("short", 2, "a", "b", "c") });
        Assert.Equal("Shareda; Sharedb; Uniquec", result.Citations[0].Content);
        Assert.Equal("Doea; Roeb; Roec", result.Citations[1].Content);
    }

    [Theory]
    [InlineData("style")]
    [InlineData("citation")]
    [InlineData("name")]
    public void SubsequentEtAlOptionsTriggerComparisonWithoutAPositionCondition(string owner) {
        const string options = "et-al-min=\"3\" et-al-use-first=\"2\" et-al-subsequent-min=\"2\" et-al-subsequent-use-first=\"1\"";
        string header = owner == "style" ? Header.Replace("class=\"note\">", "class=\"note\" " + options + ">") : Header;
        CslStyle style = CslStyle.Parse(header + "<citation disambiguate-add-names=\"true\" " + (owner == "citation" ? options : "") + "><layout delimiter=\"; \"><names variable=\"author\"><name form=\"short\" " + (owner == "name" ? options : "") + "/></names></layout></citation></style>");
        BibliographyDocument data = Data("[{\"family\":\"Doe\"},{\"family\":\"Roe\"}]", "[{\"family\":\"Doe\"},{\"family\":\"Smith\"}]");
        Assert.Equal("Doe, Roe; Doe, Smith", new CslProcessor(data, style).Render(new[] {
            Cite("first", 1, "a", "b"), Cite("short", 2, "a", "b") }).Citations.Last().Content);
    }

    [Fact]
    public void FirstNoteVariableConditionsTriggerComparisonWithoutAPositionCondition() {
        CslStyle style = CslStyle.Parse(Header + "<citation><layout delimiter=\"; \"><choose><if variable=\"first-reference-note-number\">" + Short + "<text variable=\"first-reference-note-number\" prefix=\", supra \"/></if><else>" + First + "</else></choose></layout></citation></style>");
        Assert.Equal("Doe, Alpha, supra 1; Doe, Beta, supra 1", new CslProcessor(Data(), style).Render(new[] {
            Cite("first", 1, "a", "b"), Cite("short", 2, "a", "b") }).Citations.Last().Content);
    }

    [Theory]
    [InlineData(false, "Collected Essays, Doe; Collected Essays, Roe")]
    [InlineData(true, "Collected Essaysa; Collected Essaysb")]
    public void FirstAndSubsequentFormsOfDifferentWorksCanCollide(bool yearSuffix, string expected) {
        BibliographyDocument data = BibliographyDocument.Parse("[{\"id\":\"a\",\"type\":\"book\",\"title\":\"Collected Essays\",\"title-short\":\"Essays\",\"author\":[{\"family\":\"Doe\"}]},{\"id\":\"b\",\"type\":\"book\",\"title\":\"Other Writings\",\"title-short\":\"Collected Essays\",\"author\":[{\"family\":\"Roe\"}]}]", BibliographyFormat.CslJson).Document;
        string detail = yearSuffix ? "<text variable=\"year-suffix\"/>" : "<choose><if disambiguate=\"true\"><names variable=\"author\" prefix=\", \"><name form=\"short\"/></names></if></choose>";
        CslStyle style = CslStyle.Parse(Header + "<citation " + (yearSuffix ? "disambiguate-add-year-suffix=\"true\"" : "") + "><layout delimiter=\"; \"><choose><if position=\"first\"><text variable=\"title\"/></if><else><text variable=\"title\" form=\"short\"/></else></choose>" + detail + "</layout></citation><bibliography><sort><key variable=\"title\"/></sort><layout><text variable=\"title\"/><text variable=\"year-suffix\"/></layout></bibliography></style>");
        CslRenderResult result = new CslProcessor(data, style).Render(new[] { Cite("first-b", 1, "b"), Cite("mixed", 2, "a", "b") });
        Assert.Equal(expected, result.Citations.Last().Content);
        if (yearSuffix) Assert.Equal(new[] { "Collected Essaysa", "Other Writingsb" }, result.Bibliography.Select(entry => entry.Content));
    }

    [Fact]
    public void EqualFormsOfOneWorkDoNotCreateAnAmbiguity() {
        CslStyle style = CslStyle.Parse(Header + "<citation><layout><choose><if position=\"first\"><text variable=\"title\"/></if><else><text variable=\"title\"/></else></choose><choose><if disambiguate=\"true\"><text value=\"Extra\"/></if></choose></layout></citation></style>");
        Assert.Equal(new[] { "Alpha", "Alpha" }, new CslProcessor(Data(), style).Render(new[] { Cite("first", 1, "a"), Cite("short", 2, "a") }).Citations.Select(entry => entry.Content));
    }

    [Theory]
    [InlineData("all-names")]
    [InlineData("primary-name")]
    [InlineData("by-cite")]
    public void AShortFormCannotDowngradeTheExpansionRequiredByTheFullForm(string rule) {
        CslStyle style = CslStyle.Parse(Header + "<citation disambiguate-add-givenname=\"true\" givenname-disambiguation-rule=\"" + rule + "\"><layout delimiter=\"; \"><choose><if position=\"first\"><names variable=\"author\"><name form=\"long\" initialize-with=\". \"/></names></if><else><names variable=\"author\"><name form=\"short\"/></names></else></choose></layout></citation></style>");
        CslRenderResult result = new CslProcessor(Data(secondNames: "[{\"family\":\"Doe\",\"given\":\"Jane\"}]"), style).Render(new[] {
            Cite("first", 1, "a", "b"), Cite("short", 2, "a", "b") });
        Assert.Equal(new[] { "John Doe; Jane Doe", "John Doe; Jane Doe" }, result.Citations.Select(entry => entry.Content));
    }
}
