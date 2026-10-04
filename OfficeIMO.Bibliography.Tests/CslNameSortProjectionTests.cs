namespace OfficeIMO.Bibliography.Tests;

public sealed class CslNameSortProjectionTests {
    [Theory]
    [InlineData("suffix", "III", "II")]
    [InlineData("dropping-particle", "von", "de")]
    public void ShortNameKeysOmitPartsThatTheShortFormDoesNotDisplay(string part, string first, string second) {
        string json = "[{\"id\":\"alpha\",\"type\":\"book\",\"title\":\"Alpha\",\"author\":[{\"family\":\"Smith\",\"" + part + "\":\"" + first + "\"}]}," +
            "{\"id\":\"beta\",\"type\":\"book\",\"title\":\"Beta\",\"author\":[{\"family\":\"Smith\",\"" + part + "\":\"" + second + "\"}]}]";
        Assert.Equal(new[] { "alpha", "beta" }, Render(json, Names("short"), "macro=\"authors\"").Bibliography.Select(entry => entry.Key));
    }

    [Theory]
    [InlineData("never")]
    [InlineData("sort-only")]
    [InlineData("display-and-sort")]
    public void ShortNameKeysRetainNonDroppingParticles(string demote) {
        const string json = """
            [{"id":"van","type":"book","title":"Alpha","author":[{"family":"Gogh","given":"Aaron","non-dropping-particle":"van","dropping-particle":"de","suffix":"II"}]},
             {"id":"de","type":"book","title":"Zeta","author":[{"family":"Gogh","given":"Zoe","non-dropping-particle":"de","dropping-particle":"von","suffix":"III"}]}]
            """;
        Assert.Equal(new[] { "de", "van" }, Render(json, Names("short"), "macro=\"authors\"", demote).Bibliography.Select(entry => entry.Key));
    }

    [Theory]
    [InlineData("short", "macro=\"authors\"")]
    [InlineData("long", "macro=\"authors\"")]
    [InlineData("long", "variable=\"author\"")]
    public void GivenOnlyMononymsSortByTheirVisibleName(string form, string key) {
        string family = form == "short" ? "Smith" : "Abacus";
        string json = $$"""
            [{"id":"banksy","type":"book","author":[{"given":"Banksy"}]},
             {"id":"family","type":"book","author":[{"family":"{{family}}"}]}]
            """;
        Assert.Equal(form == "short" ? new[] { "banksy", "family" } : new[] { "family", "banksy" }, Render(json, Names(form), key).Bibliography.Select(entry => entry.Key));
    }

    [Theory]
    [InlineData("family")]
    [InlineData("literal")]
    public void NameAndTextFallbackKeysWithEqualVisibleTextUseTheSecondaryKey(string part) {
        string json = "[{\"id\":\"name\",\"type\":\"book\",\"title\":\"Aardvark\",\"author\":[{\"" + part + "\":\"Agency\"}]}," +
            "{\"id\":\"text\",\"type\":\"book\",\"title\":\"Agency\"}]";
        const string macro = "<names variable=\"author\"><name/><substitute><text variable=\"title\"/></substitute></names>";
        CslRenderResult result = Render(json, macro, "macro=\"authors\"");
        Assert.Equal(new[] { "name", "text" }, result.Bibliography.Select(entry => entry.Key));
        Assert.All(result.Bibliography, entry => Assert.Equal("Agency", entry.Content));
    }

    [Fact]
    public void NameAndOrdinaryChooseBranchKeysWithEqualTextUseTheSecondaryKey() {
        const string json = """
            [{"id":"name","type":"book","title":"Aardvark","author":[{"family":"Agency"}]},
             {"id":"text","type":"book","title":"Zeta","publisher":"Agency"}]
            """;
        const string macro = "<choose><if variable=\"author\"><names variable=\"author\"><name/></names></if><else><text variable=\"publisher\"/></else></choose>";
        Assert.Equal(new[] { "name", "text" }, Render(json, macro, "macro=\"authors\"").Bibliography.Select(entry => entry.Key));
    }

    [Theory]
    [InlineData("\u00a0", "variable=\"author\"")]
    [InlineData("\u2009", "variable=\"author\"")]
    [InlineData("\u00a0", "macro=\"authors\"")]
    [InlineData("\u2009", "macro=\"authors\"")]
    public void InstitutionalArticleStrippingRecognizesPreservedUnicodeWordSeparators(string space, string key) {
        string name = "The" + space + "Alpha Institute";
        string json = "[{\"id\":\"alpha\",\"type\":\"book\",\"author\":[{\"literal\":\"" + name + "\"}]}," +
            "{\"id\":\"middle\",\"type\":\"book\",\"author\":[{\"literal\":\"Middle Institute\"}]}]";
        CslRenderResult result = Render(json, Names("long"), key);
        Assert.Equal(new[] { "alpha", "middle" }, result.Bibliography.Select(entry => entry.Key));
        Assert.Equal(name, result.Bibliography[0].Content);
    }

    private static string Names(string form) => "<names variable=\"author\"><name form=\"" + form + "\"/></names>";

    private static CslRenderResult Render(string json, string macro, string key, string demote = "display-and-sort") {
        BibliographyDocument document = BibliographyDocument.Parse(json, BibliographyFormat.CslJson).Document;
        CslStyle style = CslStyle.Parse($"""
            <style xmlns="http://purl.org/net/xbiblio/csl" version="1.0" class="in-text" demote-non-dropping-particle="{demote}">
              <macro name="authors">{macro}</macro>
              <citation><layout><text variable="title"/></layout></citation>
              <bibliography><sort><key {key}/><key variable="title"/></sort><layout><text macro="authors"/></layout></bibliography>
            </style>
            """);
        return new CslProcessor(document, style).Render(Array.Empty<CslCitation>(), includeUncitedItems: true);
    }
}
