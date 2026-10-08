using System.Text.Json;

namespace OfficeIMO.Bibliography.Tests;

public sealed class CslCompoundInitialContractTests {
    [Theory]
    [InlineData("Guo-ping", true, true, ". ", "G.-p. Chen")]
    [InlineData("Guo-ping", false, true, ". ", "G. p. Chen")]
    [InlineData("Guo-ping", true, true, "", "G-p Chen")]
    [InlineData("Guo-ping", false, true, "", "Gp Chen")]
    [InlineData("Guo-ping", true, false, ". ", "Guo-ping Chen")]
    [InlineData("Guo Ping", true, true, ". ", "G. P. Chen")]
    [InlineData("Guo de Ping", true, true, ". ", "G. de P. Chen")]
    [InlineData("Guo\u0301-ping", true, true, ". ", "G.-p. Chen")]
    [InlineData("Émile-étienne", true, true, ". ", "É.-é. Chen")]
    [InlineData("Guo\t \r\nPing", true, true, ". ", "G. P. Chen")]
    [InlineData("Guo\u2028Ping", true, true, ". ", "G. P. Chen")]
    public void HyphenatedGivenNameComponentsAreInitializedWithoutTreatingTheContinuationAsAParticle(string given, bool hyphen, bool initialize, string suffix, string expected) {
        Assert.Equal(expected, Render(given, hyphen, initialize, suffix, html: false));
    }

    [Theory]
    [InlineData(true, "<i>G.</i>-<b>p.</b> Chen")]
    [InlineData(false, "<i>G.</i> <b>p.</b> Chen")]
    public void InitializationRetainsTheFormattingAroundEachCompoundComponent(bool hyphen, string expected) {
        Assert.Equal(expected, Render("<i>Guo</i>-<b>ping</b>", hyphen, true, ". ", html: true));
    }

    private static string Render(string given, bool hyphen, bool initialize, string suffix, bool html) {
        string json = "[{\"id\":\"a\",\"type\":\"book\",\"author\":[{\"family\":\"Chen\",\"given\":" + JsonSerializer.Serialize(given) + "}]}]";
        string style = "<style xmlns=\"http://purl.org/net/xbiblio/csl\" version=\"1.0\" class=\"in-text\" initialize-with-hyphen=\"" + hyphen.ToString().ToLowerInvariant() + "\"><citation><layout><names variable=\"author\"><name initialize=\"" + initialize.ToString().ToLowerInvariant() + "\" initialize-with=\"" + suffix + "\"/></names></layout></citation></style>";
        var cluster = new CslCitation("cluster"); cluster.Items.Add(new CslCitationItem("a"));
        return new CslProcessor(BibliographyDocument.Parse(json, BibliographyFormat.CslJson).Document, CslStyle.Parse(style), new CslRenderOptions { OutputFormat = html ? CslOutputFormat.Html : CslOutputFormat.PlainText }).Render(new[] { cluster }).Citations.Single().Content;
    }
}
