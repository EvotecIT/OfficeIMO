using System.Text.Json;

namespace OfficeIMO.Bibliography.Tests;

public sealed class CslNameCountContractTests {
    private const string Header = "<style xmlns=\"http://purl.org/net/xbiblio/csl\" version=\"1.0\" class=\"in-text\">";

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MissingAuthorsUseTheInheritedCountFormForEditors(bool html) {
        const string layout = "<group delimiter=\"|\"><names variable=\"author\"><name form=\"count\"/><substitute><names variable=\"editor\"/></substitute></names><names variable=\"editor\"><name/></names></group>";
        Assert.Equal("2", Render("\"editor\":" + Names(2), layout, html: html).Citations.Single().Content);
    }

    [Fact]
    public void AnEmptyCountContinuesToTheFirstAvailableSubstitute() {
        const string layout = "<names variable=\"author\"><name form=\"count\"/><substitute><names variable=\"editor\"/><text variable=\"title\"/></substitute></names>";
        Assert.Equal("Known", Render("\"title\":\"Known\"", layout).Citations.Single().Content);
        Assert.Equal("", Render("", layout).Citations.Single().Content);
    }

    [Theory]
    [InlineData(false, "0")]
    [InlineData(true, "1")]
    public void ASelectedZeroCountDoesNotSubstituteForAPresentContributorList(bool last, string expected) {
        string layout = "<names variable=\"author\"><name form=\"count\" et-al-min=\"1\" et-al-use-first=\"0\" et-al-use-last=\"" + last.ToString().ToLowerInvariant() + "\"/><substitute><text variable=\"title\"/></substitute></names>";
        Assert.Equal(expected, Render("\"author\":" + Names(2) + ",\"title\":\"Fallback\"", layout).Citations.Single().Content);
    }

    [Fact]
    public void SortOverrideWithZeroSelectedNamesIsANumericKeyInsteadOfAMissingKey() {
        string json = "[{\"id\":\"one\",\"type\":\"book\",\"title\":\"One\",\"author\":" + Names(1) + "},{\"id\":\"zero\",\"type\":\"book\",\"title\":\"Zero\",\"author\":" + Names(2) + "}]";
        string style = Header + "<macro name=\"count\"><names variable=\"author\"><name form=\"count\"/><substitute><text variable=\"title\"/></substitute></names></macro><citation><sort><key macro=\"count\" names-min=\"2\" names-use-first=\"0\"/></sort><layout delimiter=\"|\"><text variable=\"title\"/></layout></citation></style>";
        var cluster = new CslCitation("cluster"); cluster.Items.Add(new CslCitationItem("one")); cluster.Items.Add(new CslCitationItem("zero"));
        Assert.Equal("Zero|One", new CslProcessor(BibliographyDocument.Parse(json, BibliographyFormat.CslJson).Document, CslStyle.Parse(style)).Render(new[] { cluster }).Citations.Single().Content);
    }

    [Theory]
    [InlineData("long")]
    [InlineData("short")]
    public void SelectingZeroNamesCanStillRenderEtAlWithAnInvertedNameDelimiter(string form) {
        string layout = "<names variable=\"author\"><name form=\"" + form + "\" et-al-min=\"1\" et-al-use-first=\"0\" delimiter-precedes-et-al=\"after-inverted-name\"/></names>";
        Assert.Equal("et al.", Render("\"author\":" + Names(2), layout).Citations.Single().Content);
    }

    [Theory]
    [InlineData(2, false, "2")]
    [InlineData(3, false, "1")]
    [InlineData(4, false, "1")]
    [InlineData(3, true, "2")]
    [InlineData(4, true, "2")]
    public void CountsIncludeOnlyNamesSelectedByAbbreviation(int count, bool last, string expected) =>
        Assert.Equal(expected, Render("\"author\":" + Names(count), "<names variable=\"author\"><name form=\"count\" et-al-min=\"3\" et-al-use-first=\"1\" et-al-use-last=\"" + last.ToString().ToLowerInvariant() + "\"/></names>").Citations.Single().Content);

    [Fact]
    public void SubsequentCitationCountsUseSubsequentAbbreviationAndTheLastName() {
        const string layout = "<names variable=\"author\"><name form=\"count\" et-al-min=\"3\" et-al-use-first=\"1\" et-al-subsequent-min=\"3\" et-al-subsequent-use-first=\"2\" et-al-use-last=\"true\"/></names>";
        CslRenderResult result = Render("\"author\":" + Names(4), layout, repeat: true);
        Assert.Equal(new[] { "2", "3" }, result.Citations.Select(citation => citation.Content));
    }

    [Theory]
    [InlineData(true, "2")]
    [InlineData(false, "4")]
    public void CombinedEditorAndTranslatorListsAreCountedAfterRoleCollapsing(bool same, string expected) {
        string fields = "\"editor\":" + Names(2) + ",\"translator\":" + Names(2, same ? "Person" : "Other");
        Assert.Equal(expected, Render(fields, "<names variable=\"editor translator\"><name form=\"count\"/></names>").Citations.Single().Content);
    }

    [Fact]
    public void CountsSumTheSelectedNamesAcrossContributorVariables() =>
        Assert.Equal("3", Render("\"author\":" + Names(1) + ",\"editor\":" + Names(2), "<names variable=\"author editor\"><name form=\"count\"/></names>").Citations.Single().Content);

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void CountSortKeysUseSubstitutionAndKeyLevelAbbreviationOverrides(bool descending, bool citationSort) {
        string json = "[{\"id\":\"two\",\"type\":\"book\",\"title\":\"A\",\"editor\":" + Names(2) + "}," +
            "{\"id\":\"three\",\"type\":\"book\",\"title\":\"B\",\"author\":" + Names(3) + "}," +
            "{\"id\":\"one\",\"type\":\"book\",\"title\":\"C\",\"author\":" + Names(1) + "}]";
        string sort = "<sort><key macro=\"count\" names-min=\"3\" names-use-first=\"3\" sort=\"" + (descending ? "descending" : "ascending") + "\"/><key variable=\"title\"/></sort>";
        string macro = "<macro name=\"count\"><names variable=\"author\"><name form=\"count\" et-al-min=\"3\" et-al-use-first=\"1\"/><substitute><names variable=\"editor\"/></substitute></names></macro>";
        CslStyle style = CslStyle.Parse(Header + macro + "<citation>" + (citationSort ? sort : "") + "<layout delimiter=\"|\"><text variable=\"title\"/></layout></citation><bibliography>" + sort + "<layout><text variable=\"title\"/></layout></bibliography></style>");
        BibliographyDocument document = BibliographyDocument.Parse(json, BibliographyFormat.CslJson).Document;
        var cluster = new CslCitation("cluster");
        foreach (string key in new[] { "two", "three", "one" }) cluster.Items.Add(new CslCitationItem(key));
        CslRenderResult result = new CslProcessor(document, style).Render(new[] { cluster });
        Assert.Equal(descending ? new[] { "three", "two", "one" } : new[] { "one", "two", "three" }, result.Bibliography.Select(entry => entry.Key));
        if (citationSort) Assert.Equal(descending ? "B|A|C" : "C|A|B", result.Citations.Single().Content);
    }

    [Theory]
    [InlineData(true, "C|A")]
    [InlineData(false, "A|C")]
    public void SortKeyLastNameOverrideChangesTheRenderedCount(bool useLast, string expected) {
        string json = "[{\"id\":\"four\",\"type\":\"book\",\"title\":\"A\",\"author\":" + Names(4) + "}," +
            "{\"id\":\"two\",\"type\":\"book\",\"title\":\"C\",\"author\":" + Names(2) + "}]";
        string style = Header + "<macro name=\"count\"><names variable=\"author\"><name form=\"count\" et-al-min=\"3\" et-al-use-first=\"1\" et-al-use-last=\"true\"/></names></macro><citation><sort><key macro=\"count\" names-min=\"3\" names-use-first=\"1\" names-use-last=\"" + useLast.ToString().ToLowerInvariant() + "\"/><key variable=\"title\" sort=\"descending\"/></sort><layout delimiter=\"|\"><text variable=\"title\"/></layout></citation></style>";
        var cluster = new CslCitation("cluster"); cluster.Items.Add(new CslCitationItem("four")); cluster.Items.Add(new CslCitationItem("two"));
        Assert.Equal(expected, new CslProcessor(BibliographyDocument.Parse(json, BibliographyFormat.CslJson).Document, CslStyle.Parse(style)).Render(new[] { cluster }).Citations.Single().Content);
    }

    private static string Names(int count, string prefix = "Person") => JsonSerializer.Serialize(Enumerable.Range(1, count).Select(index => new { family = prefix + index }));

    private static CslRenderResult Render(string fields, string layout, bool html = false, bool repeat = false) {
        BibliographyDocument document = BibliographyDocument.Parse("[{\"id\":\"a\",\"type\":\"book\"" + (fields.Length > 0 ? "," + fields : "") + "}]", BibliographyFormat.CslJson).Document;
        CslStyle style = CslStyle.Parse(Header + "<citation><layout>" + layout + "</layout></citation></style>");
        var first = new CslCitation("first"); first.Items.Add(new CslCitationItem("a"));
        var second = new CslCitation("second"); second.Items.Add(new CslCitationItem("a"));
        return new CslProcessor(document, style, new CslRenderOptions { OutputFormat = html ? CslOutputFormat.Html : CslOutputFormat.PlainText }).Render(repeat ? new[] { first, second } : new[] { first });
    }
}
