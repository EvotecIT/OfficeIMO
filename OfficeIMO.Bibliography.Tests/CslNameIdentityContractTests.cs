using System.Text.Json;

namespace OfficeIMO.Bibliography.Tests;

public sealed class CslNameIdentityContractTests {
    private const string Header = "<style xmlns=\"http://purl.org/net/xbiblio/csl\" version=\"1.0\" class=\"in-text\">";

    [Theory]
    [InlineData("J. J.", "J.J.")]
    [InlineData("J J", "J.J.")]
    [InlineData("J.\u00a0J.", "J.J.")]
    [InlineData("J.\tJ.", "J.J.")]
    [InlineData("J.-J.", "J-J")]
    [InlineData("Ι. Ι.", "Ι.Ι.")]
    [InlineData("J\U0001D165. J.", "J\U0001D165.J.")]
    public void TypographicalInitialVariantsDoNotInventDifferentPeople(string first, string second) {
        string body = "<citation disambiguate-add-givenname=\"true\" givenname-disambiguation-rule=\"all-names\"><layout delimiter=\"; \"><names variable=\"author\"><name form=\"short\" initialize-with=\". \"/></names><text variable=\"title\" prefix=\", \"/></layout></citation>";
        Assert.Equal("Doe, Alpha; Doe, Beta", Render(Authors(first, second), body));
    }

    [Theory]
    [InlineData("J.J.", "J. R.", "J. J. Doe, Alpha; J. R. Doe, Beta")]
    [InlineData("J.J.", "John James", "J.J. Doe, Alpha; John James Doe, Beta")]
    public void DistinctInitialsAndFullNamesRemainAvailableForDisambiguation(string first, string second, string expected) {
        string body = "<citation disambiguate-add-givenname=\"true\" givenname-disambiguation-rule=\"all-names\"><layout delimiter=\"; \"><names variable=\"author\"><name form=\"short\" initialize-with=\". \"/></names><text variable=\"title\" prefix=\", \"/></layout></citation>";
        Assert.Equal(expected, Render(Authors(first, second), body));
    }

    [Theory]
    [InlineData(false, "J. J. Doe (editor & translator)")]
    [InlineData(true, "1")]
    public void EquivalentEditorAndTranslatorInitialsShareTheContributorRole(bool count, string expected) {
        var data = new[] { new { id = "one", type = "book", editor = new[] { new { family = "Doe", given = "J. J." } }, translator = new[] { new { family = "Doe", given = "J.J." } } } };
        string body = "<citation><layout><names variable=\"editor translator\" delimiter=\"|\"><name initialize-with=\". \" form=\"" + (count ? "count" : "long") + "\"/>" + (count ? "" : "<label prefix=\" (\" suffix=\")\"/>") + "</names></layout></citation>";
        Assert.Equal(expected, Render(JsonSerializer.Serialize(data), body));
    }

    [Fact]
    public void CorporateLiteralNamesAreComparedAsLiteralText() {
        var data = new[] { new { id = "one", type = "book", editor = new[] { new { literal = "J. J." } }, translator = new[] { new { literal = "J.J." } } } };
        Assert.Equal("2", Render(JsonSerializer.Serialize(data), "<citation><layout><names variable=\"editor translator\"><name form=\"count\"/></names></layout></citation>"));
    }

    private static string Authors(string first, string second) => JsonSerializer.Serialize(new[] {
        new { id = "one", type = "book", title = "Alpha", author = new[] { new { family = "Doe", given = first } } },
        new { id = "two", type = "book", title = "Beta", author = new[] { new { family = "Doe", given = second } } }
    });

    private static string Render(string json, string body) {
        BibliographyDocument data = BibliographyDocument.Parse(json, BibliographyFormat.CslJson).Document;
        var citation = new CslCitation("cluster");
        foreach (BibliographyItem item in data.Items) citation.Items.Add(new CslCitationItem(item.Key));
        return new CslProcessor(data, CslStyle.Parse(Header + body + "</style>")).Render(new[] { citation }).Citations.Single().Content;
    }
}
