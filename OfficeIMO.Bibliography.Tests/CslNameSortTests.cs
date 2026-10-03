using System.Text.Json;

namespace OfficeIMO.Bibliography.Tests;

public sealed class CslNameSortTests {
    [Theory]
    [InlineData("display-and-sort", false)]
    [InlineData("display-and-sort", true)]
    [InlineData("sort-only", false)]
    [InlineData("sort-only", true)]
    public void DemotedParticlesSortBeforeGivenNames(string demote, bool macro) {
        var names = new[] {
            new { id = "van", type = "book", author = new[] { new { family = "Gogh", given = "Aaron", particle = "van" } } },
            new { id = "de", type = "book", author = new[] { new { family = "Gogh", given = "Zoe", particle = "de" } } }
        };
        string json = JsonSerializer.Serialize(names).Replace("\"particle\"", "\"non-dropping-particle\"");
        CslRenderResult result = Render(json, demote, macro);
        Assert.Equal(new[] { "de", "van" }, result.Bibliography.Select(entry => entry.Key));
        Assert.Equal(new[] { "Zoe de Gogh", "Aaron van Gogh" }, result.Bibliography.Select(entry => entry.Content));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DroppingParticlesSortBeforeGivenNamesWhenNonDroppingParticlesStayWithTheFamily(bool macro) {
        const string json = """
            [{"id":"von","type":"book","author":[{"family":"Gogh","given":"Aaron","non-dropping-particle":"van","dropping-particle":"von"}]},
             {"id":"de","type":"book","author":[{"family":"Gogh","given":"Zoe","non-dropping-particle":"van","dropping-particle":"de"}]}]
            """;
        Assert.Equal(new[] { "de", "von" }, Render(json, "never", macro).Bibliography.Select(entry => entry.Key));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MissingParticlesRetainTheirSortPositionBeforeGivenNames(bool macro) {
        const string json = """
            [{"id":"particle","type":"book","author":[{"family":"Gogh","given":"Aaron","non-dropping-particle":"van"}]},
             {"id":"plain","type":"book","author":[{"family":"Gogh","given":"Zoe"}]}]
            """;
        Assert.Equal(new[] { "plain", "particle" }, Render(json, "display-and-sort", macro).Bibliography.Select(entry => entry.Key));
    }

    [Theory]
    [InlineData("The", false)]
    [InlineData("An", false)]
    [InlineData("A", false)]
    [InlineData("the", true)]
    [InlineData("an", true)]
    [InlineData("a", true)]
    public void InstitutionalArticlesAreRemovedOnlyFromSortKeys(string article, bool macro) {
        string name = article + " Zeta Institute";
        string json = "[{\"id\":\"zeta\",\"type\":\"book\",\"author\":[{\"literal\":" + JsonSerializer.Serialize(name) + "}]}," +
            "{\"id\":\"middle\",\"type\":\"book\",\"author\":[{\"literal\":\"Middle Institute\"}]}]";
        CslRenderResult result = Render(json, "display-and-sort", macro);
        Assert.Equal(new[] { "middle", "zeta" }, result.Bibliography.Select(entry => entry.Key));
        Assert.Equal(name, result.Bibliography.Last().Content);
    }

    [Fact]
    public void SortsCompareEachNameInOrderWithoutTheDisplayedConjunction() {
        const string json = """
            [{"id":"two","type":"book","author":[{"family":"Doe"},{"family":"Zeta"}]},
             {"id":"three","type":"book","author":[{"family":"Doe"},{"family":"Alpha"},{"family":"Zeta"}]}]
            """;
        Assert.Equal(new[] { "three", "two" }, Render(json, "display-and-sort", true, "and=\"text\"").Bibliography.Select(entry => entry.Key));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LiteralAndFamilyOnlyNamesWithTheSameTextShareSecondarySortKeys(bool macro) {
        const string json = """
            [{"id":"beta","type":"book","title":"Beta","author":[{"literal":"Agency"}]},
             {"id":"alpha","type":"book","title":"Alpha","author":[{"family":"Agency"}]}]
            """;
        Assert.Equal(new[] { "alpha", "beta" }, Render(json, "display-and-sort", macro).Bibliography.Select(entry => entry.Key));
    }

    private static CslRenderResult Render(string json, string demote, bool macro, string nameOptions = "") {
        BibliographyDocument document = BibliographyDocument.Parse(json, BibliographyFormat.CslJson).Document;
        string key = macro ? "macro=\"authors\"" : "variable=\"author\"";
        CslStyle style = CslStyle.Parse($"""
            <style xmlns="http://purl.org/net/xbiblio/csl" version="1.0" class="in-text" demote-non-dropping-particle="{demote}">
              <macro name="authors"><names variable="author"><name {nameOptions}/></names></macro>
              <citation><layout><text variable="title"/></layout></citation>
              <bibliography><sort><key {key}/><key variable="title"/></sort><layout><names variable="author"><name/></names></layout></bibliography>
            </style>
            """);
        return new CslProcessor(document, style).Render(Array.Empty<CslCitation>(), includeUncitedItems: true);
    }
}
