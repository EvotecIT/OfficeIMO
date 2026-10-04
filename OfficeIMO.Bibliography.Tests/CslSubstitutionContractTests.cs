namespace OfficeIMO.Bibliography.Tests;

public sealed class CslSubstitutionContractTests {
    private const string Header = "<style xmlns=\"http://purl.org/net/xbiblio/csl\" version=\"1.0\" class=\"in-text\">";

    [Theory]
    [InlineData("\"title\":\"Known\"", "<text variable=\"title\"/>", "Known")]
    [InlineData("\"volume\":\"2\"", "<number variable=\"volume\"/>", "2")]
    [InlineData("\"issued\":{\"date-parts\":[[2024]]}", "<date variable=\"issued\"><date-part name=\"year\"/></date>", "2024")]
    public void SubstituteMacrosSuppressRepeatedVariablesAsTheyRender(string fields, string rendering, string expected) {
        string macro = "<macro name=\"replacement\">" + rendering + rendering + "</macro>";
        Assert.Equal(expected, Render(fields, "<names variable=\"author\"><substitute><text macro=\"replacement\"/></substitute></names>" + rendering, macro));
    }

    [Fact]
    public void RepeatedEditorNamesInsideASubstituteMacroRenderOnce() {
        const string names = "<names variable=\"editor\"><name/><label prefix=\" \"/></names>";
        string macro = "<macro name=\"editor\">" + names + names + "</macro>";
        Assert.Equal("Editor editor", Render("\"editor\":[{\"literal\":\"Editor\"}]",
            "<names variable=\"author\"><substitute><text macro=\"editor\"/></substitute></names>" + names, macro));
    }

    [Fact]
    public void VariablesRenderedBeforeSubstitutionRemainAvailable() {
        const string names = "<names variable=\"editor\"><name/></names>";
        Assert.Equal("Editor|Editor|Editor", Render("\"editor\":[{\"literal\":\"Editor\"}]",
            "<group delimiter=\"|\">" + names + "<names variable=\"author\"><substitute>" + names + "</substitute></names>" + names + "</group>"));
    }

    [Fact]
    public void ALabeledVariableIsNotTreatedAsPreviouslyRenderedData() =>
        Assert.Equal("volume|2", Render("\"volume\":\"2\"",
            "<group delimiter=\"|\"><label variable=\"volume\"/><names variable=\"author\"><substitute><text variable=\"volume\"/></substitute></names><text variable=\"volume\"/></group>"));

    [Fact]
    public void EmptyFormattedFallbacksDoNotSuppressLaterCandidates() =>
        Assert.Equal(".", Render("\"title\":\".\"",
            "<names variable=\"author\"><substitute><text macro=\"title\" strip-periods=\"true\"/><text variable=\"title\"/></substitute></names><text variable=\"title\" prefix=\" repeated=\"/>",
            "<macro name=\"title\"><text variable=\"title\"/></macro>"));

    [Theory]
    [InlineData("<text macro=\"fallback\"/>")]
    [InlineData("<group><text macro=\"fallback\"/><text variable=\"URL\"/></group>")]
    public void MacrosSuppressTermsWhoseAccompanyingVariableIsEmpty(string layout) =>
        Assert.Equal("", Render("", layout, "<macro name=\"fallback\"><text value=\"Unknown\"/><text variable=\"title\"/></macro>"));

    [Fact]
    public void ExplicitGroupsInsideMacrosStillSuppressEmptyVariables() =>
        Assert.Equal("", Render("", "<text macro=\"fallback\"/>",
            "<macro name=\"fallback\"><group><text value=\"Unknown\"/><text variable=\"title\"/></group></macro>"));

    [Fact]
    public void CombinedContributorSubstitutionSuppressesItsOriginalRoles() {
        const string fields = "\"editor\":[{\"literal\":\"Editor\"}],\"translator\":[{\"literal\":\"Editor\"}]";
        Assert.Equal("Editor", Render(fields,
            "<names variable=\"author\"><substitute><names variable=\"editor translator\"><name/></names></substitute></names><names variable=\"editor\"><name/></names><names variable=\"translator\"><name/></names>"));
    }

    [Fact]
    public void DiscardedNameCandidatesDoNotBecomeTheBibliographyComparisonNames() {
        const string json = "[{\"id\":\"a\",\"type\":\"book\",\"editor\":[{\"literal\":\".\"}],\"title\":\"Alpha\"},{\"id\":\"b\",\"type\":\"book\",\"editor\":[{\"literal\":\".\"}],\"title\":\"Beta\"}]";
        string names = "<names variable=\"author\"><substitute><names variable=\"editor\" strip-periods=\"true\"><name/></names><text variable=\"title\"/></substitute></names>";
        CslStyle style = CslStyle.Parse(Header + "<citation><layout><text variable=\"id\"/></layout></citation><bibliography subsequent-author-substitute=\"—\"><layout>" + names + "</layout></bibliography></style>");
        var processor = new CslProcessor(BibliographyDocument.Parse(json, BibliographyFormat.CslJson).Document, style);
        Assert.Equal(new[] { "Alpha", "Beta" }, processor.RenderBibliography().Select(entry => entry.Content));
    }

    [Fact]
    public void DiscardedDisplayCandidatesDoNotIncreaseTheAlignmentHint() {
        string json = "[{\"id\":\"a\",\"type\":\"book\",\"abstract\":\"" + new string('.', 100) + "\",\"title\":\"Alpha\"}]";
        string names = "<names variable=\"author\"><substitute><text macro=\"discarded\" strip-periods=\"true\"/><text variable=\"title\"/></substitute></names>";
        CslStyle style = CslStyle.Parse(Header + "<macro name=\"discarded\"><text variable=\"abstract\" display=\"left-margin\"/></macro><citation><layout><text variable=\"id\"/></layout></citation><bibliography><layout>" + names + "</layout></bibliography></style>");
        var processor = new CslProcessor(BibliographyDocument.Parse(json, BibliographyFormat.CslJson).Document, style);
        CslRenderResult result = processor.Render(Array.Empty<CslCitation>(), true);
        Assert.Equal("Alpha", result.Bibliography.Single().Content);
        Assert.Equal(0, result.BibliographyLayout!.MaximumLeftMarginCharacters);
    }

    [Theory]
    [InlineData("complete-all")]
    [InlineData("complete-each")]
    [InlineData("partial-each")]
    [InlineData("partial-first")]
    public void RepeatedTextFallbacksUseTheBibliographySubstitutionRule(string rule) {
        const string json = "[{\"id\":\"a\",\"type\":\"book\",\"title\":\"Alpha\"},{\"id\":\"b\",\"type\":\"book\",\"title\":\"Alpha\"},{\"id\":\"c\",\"type\":\"book\",\"title\":\"Beta\"}]";
        string names = "<names variable=\"author\" prefix=\"[\" suffix=\"]\"><substitute><text variable=\"title\"/></substitute></names>";
        CslStyle style = CslStyle.Parse(Header + "<citation><layout><text variable=\"id\"/></layout></citation><bibliography subsequent-author-substitute=\"—\" subsequent-author-substitute-rule=\"" + rule + "\"><layout>" + names + "</layout></bibliography></style>");
        var processor = new CslProcessor(BibliographyDocument.Parse(json, BibliographyFormat.CslJson).Document, style);
        Assert.Equal(new[] { "[Alpha]", "[—]", "[Beta]" }, processor.RenderBibliography().Select(entry => entry.Content));
    }

    [Theory]
    [InlineData("editortranslator")]
    [InlineData("editor-translator")]
    public void CombinedContributorLabelsRespectLocalOverrides(string term) {
        const string fields = "\"editor\":[{\"literal\":\"Editor\"}],\"translator\":[{\"literal\":\"Editor\"}]";
        string locale = "<locale><terms><term name=\"" + term + "\">joint role</term></terms></locale>";
        Assert.Equal("Editor (joint role)", Render(fields,
            "<names variable=\"editor translator\"><name/><label prefix=\" (\" suffix=\")\"/></names>", locale));
    }

    [Theory]
    [InlineData("title")]
    [InlineData("container-title")]
    public void ShortFormsCannotResurrectASuppressedVariable(string variable) {
        string rendering = "<text variable=\"" + variable + "\" form=\"short\"/>";
        string fields = "\"" + variable + "\":\"Full\",\"" + variable + "-short\":\"Short\"";
        Assert.Equal("Short", Render(fields,
            "<names variable=\"author\"><substitute><text macro=\"replacement\"/></substitute></names>" + rendering,
            "<macro name=\"replacement\">" + rendering + rendering + "</macro>"));
    }

    [Theory]
    [InlineData(1, "complete-all")]
    [InlineData(2, "complete-each")]
    [InlineData(3, "partial-each")]
    [InlineData(3, "partial-first")]
    public void EmptyBibliographyReplacementsKeepNestedFallbacksSelected(int depth, string rule) {
        const string json = "[{\"id\":\"a\",\"type\":\"book\",\"title\":\"Alpha\"},{\"id\":\"b\",\"type\":\"book\",\"title\":\"Alpha\"},{\"id\":\"c\",\"type\":\"book\",\"title\":\"Alpha\"}]";
        string names = "<text variable=\"title\"/>";
        for (int index = 0; index <= depth; index++)
            names = "<names variable=\"author\"><substitute>" + names + "<text value=\"Unexpected fallback\"/></substitute></names>";
        CslStyle style = CslStyle.Parse(Header + "<citation><layout><text variable=\"id\"/></layout></citation><bibliography subsequent-author-substitute=\"\" subsequent-author-substitute-rule=\"" + rule + "\"><layout>" + names + "<text variable=\"title\" prefix=\" repeated=\"/></layout></bibliography></style>");
        var processor = new CslProcessor(BibliographyDocument.Parse(json, BibliographyFormat.CslJson).Document, style);
        Assert.Equal(new[] { "Alpha", "", "" }, processor.RenderBibliography().Select(entry => entry.Content));
    }

    [Theory]
    [InlineData("complete-all")]
    [InlineData("complete-each")]
    [InlineData("partial-each")]
    [InlineData("partial-first")]
    public void EmptyBibliographyNameReplacementsKeepTheirRolesSuppressed(string rule) {
        const string json = "[{\"id\":\"a\",\"type\":\"book\",\"editor\":[{\"literal\":\"Editor\"}]},{\"id\":\"b\",\"type\":\"book\",\"editor\":[{\"literal\":\"Editor\"}]},{\"id\":\"c\",\"type\":\"book\",\"editor\":[{\"literal\":\"Editor\"}]}]";
        string names = "<names variable=\"author\"><substitute><names variable=\"editor\"><name/></names><text value=\"Unexpected fallback\"/></substitute></names><names variable=\"editor\"><name/></names>";
        CslStyle style = CslStyle.Parse(Header + "<citation><layout><text variable=\"id\"/></layout></citation><bibliography subsequent-author-substitute=\"\" subsequent-author-substitute-rule=\"" + rule + "\"><layout>" + names + "</layout></bibliography></style>");
        var processor = new CslProcessor(BibliographyDocument.Parse(json, BibliographyFormat.CslJson).Document, style);
        Assert.Equal(new[] { "Editor", "", "" }, processor.RenderBibliography().Select(entry => entry.Content));
    }

    private static string Render(string fields, string layout, string macros = "") {
        BibliographyDocument document = BibliographyDocument.Parse("[{\"id\":\"one\",\"type\":\"book\"" + (fields.Length == 0 ? "" : "," + fields) + "}]", BibliographyFormat.CslJson).Document;
        CslStyle style = CslStyle.Parse(Header + macros + "<citation><layout>" + layout + "</layout></citation></style>");
        var citation = new CslCitation("citation");
        citation.Items.Add(new CslCitationItem("one"));
        return new CslProcessor(document, style).Render(new[] { citation }).Citations.Single().Content;
    }
}
