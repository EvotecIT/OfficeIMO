namespace OfficeIMO.Bibliography.Tests;

public sealed class CslContributorLabelContractTests {
    private const string Header = "<style xmlns=\"http://purl.org/net/xbiblio/csl\" version=\"1.0\" class=\"in-text\"><citation><layout><text variable=\"title\"/></layout></citation>";

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void CompleteNameReplacementPreservesTheContributorRoleAndItsOrder(bool labelFirst, bool html) {
        const string json = """
            [{"id":"a","type":"book","editor":[{"literal":"Editor"}]},
             {"id":"b","type":"book","editor":[{"literal":"Editor"}]}]
            """;
        string label = labelFirst ? "<label form=\"short\" suffix=\" \" font-weight=\"bold\"/>" : "<label form=\"short\" prefix=\" \" font-weight=\"bold\"/>";
        string names = "<names variable=\"editor\" prefix=\"[\" suffix=\"]\">" + (labelFirst ? label : "") + "<name/>" + (labelFirst ? "" : label) + "</names>";
        CslRenderResult result = Render(json, names, "—", html);
        string role = html ? "<b>ed.</b>" : "ed.";
        string first = labelFirst ? "[" + role + " Editor]" : "[Editor " + role + "]";
        string second = labelFirst ? "[" + role + " —]" : "[— " + role + "]";
        Assert.Equal(new[] { Entry(first, html), Entry(second, html) }, result.Bibliography.Select(entry => entry.Content));
    }

    [Theory]
    [InlineData(false, "—")]
    [InlineData(true, "—")]
    [InlineData(false, "")]
    [InlineData(true, "")]
    public void CompleteReplacementRetainsNamesAffixesAndRemovesNameAffixes(bool html, string replacement) {
        const string json = """
            [{"id":"a","type":"book","editor":[{"literal":"Person"}]},
             {"id":"b","type":"book","editor":[{"literal":"Person"}]}]
            """;
        const string names = "<names variable=\"editor\" prefix=\"[\" suffix=\"]\"><name prefix=\"(\" suffix=\")\"/><label form=\"short\" prefix=\" \"/></names>";
        Assert.Equal(new[] { Entry("[(Person) ed.]", html), Entry("[" + replacement + " ed.]", html) },
            Render(json, names, replacement, html).Bibliography.Select(entry => entry.Content));
    }

    [Theory]
    [InlineData("<name and=\"symbol\"/>", "A & B eds.")]
    [InlineData("<name et-al-min=\"2\" et-al-use-first=\"1\"/>", "A et al. eds.")]
    public void CompleteReplacementRemovesTheNameListTermsAndRetainsThePluralRole(string name, string first) {
        const string json = """
            [{"id":"a","type":"book","editor":[{"literal":"A"},{"literal":"B"}]},
             {"id":"b","type":"book","editor":[{"literal":"A"},{"literal":"B"}]}]
            """;
        string names = "<names variable=\"editor\">" + name + "<label form=\"short\" prefix=\" \"/></names>";
        Assert.Equal(new[] { first, "— eds." }, Render(json, names, "—").Bibliography.Select(entry => entry.Content));
    }

    [Fact]
    public void AnInheritedEditorLabelSurvivesAnAuthorToEditorReplacement() {
        const string json = """
            [{"id":"a","type":"book","author":[{"literal":"Person"}]},
             {"id":"b","type":"book","editor":[{"literal":"Person"}]}]
            """;
        const string names = "<names variable=\"author\"><name/><label form=\"short\" prefix=\" \"/><substitute><names variable=\"editor\"/></substitute></names>";
        Assert.Equal(new[] { "Person", "— ed." }, Render(json, names, "—").Bibliography.Select(entry => entry.Content));
    }

    [Fact]
    public void EmptyNameReplacementRetainsItsRoleAndSuppressesTheUsedFallbackVariable() {
        const string json = """
            [{"id":"a","type":"book","editor":[{"literal":"Person"}]},
             {"id":"b","type":"book","editor":[{"literal":"Person"}]}]
            """;
        const string names = "<names variable=\"author\"><substitute><names variable=\"editor\"><name/><label form=\"short\" prefix=\" \"/></names><text value=\"Wrong fallback\"/></substitute></names><names variable=\"editor\"><name prefix=\"Repeated=\"/></names>";
        Assert.Equal(new[] { "Person ed.", " ed." }, Render(json, names, "").Bibliography.Select(entry => entry.Content));
    }

    [Fact]
    public void ALabelOnALaterNamesExpressionRetainsTheUnreplacedNames() {
        const string json = """
            [{"id":"a","type":"book","author":[{"literal":"Author"}],"editor":[{"literal":"Editor"}]},
             {"id":"b","type":"book","author":[{"literal":"Author"}],"editor":[{"literal":"Editor"}]}]
            """;
        const string names = "<group delimiter=\"|\"><names variable=\"author\"><name/></names><names variable=\"editor\"><name/><label form=\"short\" prefix=\" \"/></names></group>";
        Assert.Equal(new[] { "Author|Editor ed.", "—|Editor ed." }, Render(json, names, "—").Bibliography.Select(entry => entry.Content));
    }

    private static string Entry(string value, bool html) => html ? "<div class=\"csl-entry\">" + value + "</div>" : value;

    private static CslRenderResult Render(string json, string names, string replacement, bool html = false) {
        CslStyle style = CslStyle.Parse(Header + "<bibliography subsequent-author-substitute=\"" + replacement + "\"><layout>" + names + "</layout></bibliography></style>");
        return new CslProcessor(BibliographyDocument.Parse(json, BibliographyFormat.CslJson).Document, style, new CslRenderOptions { OutputFormat = html ? CslOutputFormat.Html : CslOutputFormat.PlainText }).Render(Array.Empty<CslCitation>(), includeUncitedItems: true);
    }
}
