using System.Net;
using System.Text.Json;

namespace OfficeIMO.Bibliography.Tests;

public sealed class CslShortFieldAliasContractTests {
    private const string Header = "<style xmlns=\"http://purl.org/net/xbiblio/csl\" version=\"1.0\" class=\"in-text\">";

    [Theory]
    [InlineData("title", "shortTitle", false)]
    [InlineData("title", "shortTitle", true)]
    [InlineData("container-title", "journalAbbreviation", false)]
    [InlineData("container-title", "journalAbbreviation", true)]
    public void InputAliasesSupplyShortTextAndVariableConditionsWithoutRewritingSource(string variable, string alias, bool html) {
        string json = "[{\"id\":\"a\",\"type\":\"book\",\"" + variable + "\":\"Full\",\"" + alias + "\":\"Brief\"}]";
        BibliographyDocument data = BibliographyDocument.Parse(json, BibliographyFormat.CslJson).Document;
        string fields = "<group delimiter=\"|\"><text variable=\"" + variable + "-short\"/><text variable=\"" + variable + "\" form=\"short\"/><text variable=\"" + variable + "\"/><choose><if variable=\"" + variable + "-short\"><text value=\"present\"/></if><else><text value=\"missing\"/></else></choose></group>";
        Assert.Equal("Brief|Brief|Full|present", WebUtility.HtmlDecode(Citation(data, fields, html)));
        using JsonDocument saved = JsonDocument.Parse(data.Write(new BibliographyWriteOptions { Format = BibliographyFormat.CslJson, Mode = BibliographyWriterMode.Canonical }).Content);
        JsonElement item = saved.RootElement[0];
        Assert.Equal("Brief", item.GetProperty(alias).GetString());
        Assert.False(item.TryGetProperty(variable + "-short", out _));
    }

    [Theory]
    [InlineData("title", "shortTitle", "\"Preferred\"", "Preferred|Preferred|present")]
    [InlineData("container-title", "journalAbbreviation", "\"Preferred\"", "Preferred|Preferred|present")]
    [InlineData("title", "shortTitle", "\"\"", "Full|missing")]
    [InlineData("container-title", "journalAbbreviation", "\"\"", "Full|missing")]
    [InlineData("title", "shortTitle", "null", "Full|missing")]
    [InlineData("container-title", "journalAbbreviation", "null", "Full|missing")]
    public void AnExplicitCanonicalFieldTakesPrecedenceIncludingAnEmptyValue(string variable, string alias, string canonical, string expected) {
        string json = "[{\"id\":\"a\",\"type\":\"book\",\"" + alias + "\":\"Alias\",\"" + variable + "\":\"Full\",\"" + variable + "-short\":" + canonical + "}]";
        string fields = "<group delimiter=\"|\"><text variable=\"" + variable + "-short\"/><text variable=\"" + variable + "\" form=\"short\"/><choose><if variable=\"" + variable + "-short\"><text value=\"present\"/></if><else><text value=\"missing\"/></else></choose></group>";
        Assert.Equal(expected, Citation(BibliographyDocument.Parse(json, BibliographyFormat.CslJson).Document, fields));
    }

    [Theory]
    [InlineData("title", "shortTitle")]
    [InlineData("container-title", "journalAbbreviation")]
    public void BibliographySortKeysUseTheSameAliasAsRenderedShortFields(string variable, string alias) {
        string json = "[{\"id\":\"a\",\"type\":\"book\",\"" + variable + "\":\"Alpha\",\"" + alias + "\":\"Zulu\"}," +
            "{\"id\":\"b\",\"type\":\"book\",\"" + variable + "\":\"Zulu\",\"" + alias + "\":\"Alpha\"}]";
        var style = CslStyle.Parse(Header + "<citation><layout><text variable=\"title\"/></layout></citation><bibliography><sort><key variable=\"" + variable + "-short\"/></sort><layout><text variable=\"" + variable + "\" form=\"short\"/></layout></bibliography></style>");
        var entries = new CslProcessor(BibliographyDocument.Parse(json, BibliographyFormat.CslJson).Document, style).RenderBibliography();
        Assert.Equal(new[] { "b", "a" }, entries.Select(entry => entry.Key));
        Assert.Equal(new[] { "Alpha", "Zulu" }, entries.Select(entry => entry.Content));
    }

    [Theory]
    [InlineData("title", "shortTitle")]
    [InlineData("container-title", "journalAbbreviation")]
    public void AliasesCannotResurrectShortFormsSuppressedByNameSubstitution(string variable, string alias) {
        string json = "[{\"id\":\"a\",\"type\":\"book\",\"" + variable + "\":\"Full\",\"" + alias + "\":\"Brief\"}]";
        string shortText = "<text variable=\"" + variable + "\" form=\"short\"/>";
        string fields = "<names variable=\"author\"><substitute>" + shortText + "</substitute></names>" + shortText;
        Assert.Equal("Brief", Citation(BibliographyDocument.Parse(json, BibliographyFormat.CslJson).Document, fields));
    }

    private static string Citation(BibliographyDocument data, string fields, bool html = false) {
        var cite = new CslCitation("cluster");
        cite.Items.Add(new CslCitationItem("a"));
        var style = CslStyle.Parse(Header + "<citation><layout>" + fields + "</layout></citation></style>");
        return new CslProcessor(data, style, new CslRenderOptions { OutputFormat = html ? CslOutputFormat.Html : CslOutputFormat.PlainText })
            .Render(new[] { cite }).Citations.Single().Content;
    }
}
