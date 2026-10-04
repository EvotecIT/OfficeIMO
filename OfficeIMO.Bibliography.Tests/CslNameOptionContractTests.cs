namespace OfficeIMO.Bibliography.Tests;

public sealed class CslNameOptionContractTests {
    private const string Header = "<style xmlns=\"http://purl.org/net/xbiblio/csl\" version=\"1.0\" class=\"in-text\"";
    private const string Data = "[{\"id\":\"a\",\"type\":\"book\",\"author\":[{\"family\":\"A\"},{\"family\":\"B\"},{\"family\":\"C\"}]}]";

    [Theory]
    [InlineData("", "et-al-min=\"3\"", "A, B, C")]
    [InlineData("", "et-al-use-first=\"1\"", "A, B, C")]
    [InlineData("", "et-al-min=\"3\" et-al-use-first=\"1\"", "A et al.")]
    [InlineData("et-al-use-first=\"1\"", "et-al-min=\"3\"", "A et al.")]
    [InlineData("et-al-min=\"3\"", "et-al-use-first=\"2\"", "A, B, et al.")]
    public void AbbreviationRequiresBothEffectiveOptions(string rootOptions, string nameOptions, string expected) {
        string names = "<names variable=\"author\"><name " + nameOptions + "/></names>";
        CslRenderResult result = Render(Data, Header + " " + rootOptions + "><citation><layout>" + names + "</layout></citation><bibliography><layout>" + names + "</layout></bibliography></style>");
        Assert.Equal(expected, result.Citations.Single().Content);
        Assert.Equal(expected, result.Bibliography.Single().Content);
    }

    [Theory]
    [InlineData("et-al-min=\"3\"", "3")]
    [InlineData("et-al-use-first=\"1\"", "3")]
    [InlineData("et-al-min=\"3\" et-al-use-first=\"1\"", "1")]
    public void CountFormUsesTheSameEffectiveOptionPair(string nameOptions, string expected) {
        string names = "<names variable=\"author\"><name form=\"count\" " + nameOptions + "/></names>";
        CslRenderResult result = Render(Data, Header + "><citation><layout>" + names + "</layout></citation><bibliography><layout>" + names + "</layout></bibliography></style>");
        Assert.Equal(expected, result.Citations.Single().Content);
        Assert.Equal(expected, result.Bibliography.Single().Content);
    }

    [Theory]
    [InlineData("et-al-subsequent-min=\"3\"", "A, B, C", "A, B, C")]
    [InlineData("et-al-subsequent-use-first=\"1\"", "A, B, C", "A, B, C")]
    [InlineData("et-al-subsequent-min=\"3\" et-al-subsequent-use-first=\"1\"", "A, B, C", "A et al.")]
    [InlineData("et-al-min=\"3\" et-al-subsequent-use-first=\"1\"", "A, B, C", "A et al.")]
    [InlineData("et-al-use-first=\"1\" et-al-subsequent-min=\"3\"", "A, B, C", "A et al.")]
    [InlineData("et-al-min=\"3\" et-al-use-first=\"1\" et-al-subsequent-use-first=\"2\"", "A et al.", "A, B, et al.")]
    public void SubsequentOptionsReplaceOnlyTheirCorrespondingEffectiveValues(string options, string first, string second) {
        string style = Header + "><citation><layout><names variable=\"author\"><name " + options + "/></names></layout></citation></style>";
        Assert.Equal(new[] { first, second }, Render(Data, style, repeated: true).Citations.Select(entry => entry.Content));
    }

    [Fact]
    public void CitationOptionsDoNotCompleteAnUnpairedBibliographySetting() {
        const string names = "<names variable=\"author\"><name et-al-min=\"3\"/></names>";
        CslRenderResult result = Render(Data, Header + "><citation et-al-use-first=\"2\"><layout>" + names + "</layout></citation><bibliography><layout>" + names + "</layout></bibliography></style>");
        Assert.Equal("A, B, et al.", result.Citations.Single().Content);
        Assert.Equal("A, B, C", result.Bibliography.Single().Content);
    }

    [Theory]
    [InlineData("et-al-min=\"2\"", "", "One", "Two", "Three")]
    [InlineData("", "names-min=\"2\"", "One", "Two", "Three")]
    [InlineData("et-al-min=\"2\"", "names-use-first=\"1\"", "One", "Three", "Two")]
    [InlineData("et-al-use-first=\"1\"", "names-min=\"2\"", "One", "Three", "Two")]
    public void MacroSortOverridesCompleteThePairWithoutInventingAMissingValue(string nameOptions, string keyOptions, string first, string second, string third) {
        const string json = "[{\"id\":\"a\",\"type\":\"book\",\"title\":\"One\",\"author\":[{\"family\":\"A\"}]},{\"id\":\"b\",\"type\":\"book\",\"title\":\"Three\",\"author\":[{\"family\":\"A\"},{\"family\":\"B\"},{\"family\":\"C\"}]},{\"id\":\"c\",\"type\":\"book\",\"title\":\"Two\",\"author\":[{\"family\":\"A\"},{\"family\":\"B\"}]}]";
        string style = Header + "><macro name=\"count\"><names variable=\"author\"><name form=\"count\" " + nameOptions + "/></names></macro><citation><layout><text variable=\"title\"/></layout></citation><bibliography><sort><key macro=\"count\" " + keyOptions + "/></sort><layout><text variable=\"title\"/></layout></bibliography></style>";
        Assert.Equal(new[] { first, second, third }, Render(json, style).Bibliography.Select(entry => entry.Content));
    }

    [Theory]
    [InlineData(false, "Person|Person ed.|Person trans.")]
    [InlineData(true, "3")]
    public void AThirdContributorRoleIsRenderedIndependentlyOfMatchingEditorsAndTranslators(bool count, string expected) {
        const string json = "[{\"id\":\"a\",\"type\":\"book\",\"author\":[{\"literal\":\"Person\"}],\"editor\":[{\"literal\":\"Person\"}],\"translator\":[{\"literal\":\"Person\"}]}]";
        string style = Header + "><citation><layout><names variable=\"author editor translator\" delimiter=\"|\"><name form=\"" + (count ? "count" : "long") + "\"/><label form=\"short\" prefix=\" \"/></names></layout></citation></style>";
        Assert.Equal(expected, Render(json, style).Citations.Single().Content);
    }

    [Theory]
    [InlineData(false, "Person|Person ed.|Person trans.")]
    [InlineData(true, "Person|2")]
    public void SuppressingTheThirdSelectedRoleDoesNotTurnTheExpressionIntoACombinedRole(bool count, string expected) {
        const string json = "[{\"id\":\"a\",\"type\":\"book\",\"author\":[{\"literal\":\"Person\"}],\"editor\":[{\"literal\":\"Person\"}],\"translator\":[{\"literal\":\"Person\"}]}]";
        string style = Header + "><citation><layout><group delimiter=\"|\"><names variable=\"collection-editor\"><substitute><names variable=\"author\"/></substitute></names><names variable=\"author editor translator\" delimiter=\"|\"><name form=\"" + (count ? "count" : "long") + "\"/><label form=\"short\" prefix=\" \"/></names></group></layout></citation></style>";
        Assert.Equal(expected, Render(json, style).Citations.Single().Content);
    }

    private static CslRenderResult Render(string json, string style, bool repeated = false) {
        var first = new CslCitation("first"); first.Items.Add(new CslCitationItem("a"));
        var second = new CslCitation("second"); second.Items.Add(new CslCitationItem("a"));
        return new CslProcessor(BibliographyDocument.Parse(json, BibliographyFormat.CslJson).Document, CslStyle.Parse(style)).Render(repeated ? new[] { first, second } : new[] { first }, includeUncitedItems: true);
    }
}
