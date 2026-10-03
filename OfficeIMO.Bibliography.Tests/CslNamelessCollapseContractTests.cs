using System.Collections.Generic;
using System.Text.Json;

namespace OfficeIMO.Bibliography.Tests;

public sealed class CslNamelessCollapseContractTests {
    [Theory]
    [InlineData("year-suffix", 2, false, "2020a,b")]
    [InlineData("year-suffix", 3, true, "2020a,b,c")]
    [InlineData("year-suffix-ranged", 2, true, "2020a,b")]
    [InlineData("year-suffix-ranged", 3, false, "2020a–c")]
    public void DateOnlyCitesShareTheirEmptyVisibleNameGroup(string mode, int count, bool emptyNameExpression, string expected) {
        Assert.Equal(expected, Render(mode, count, emptyNameExpression, prefixIndex: -1));
    }

    [Theory]
    [InlineData(0, "see 2020a, 2020b,c")]
    [InlineData(1, "2020a, see 2020b, 2020c")]
    public void LiteralCitationPrefixesKeepTheAffixedCiteAndItsNeighborUncollapsed(int prefixIndex, string expected) {
        Assert.Equal(expected, Render("year-suffix", 3, false, prefixIndex));
    }

    private static string Render(string mode, int count, bool emptyNameExpression, int prefixIndex) {
        var data = BibliographyDocument.Parse(JsonSerializer.Serialize(Enumerable.Range(0, count).Select(index =>
            new { id = "item-" + index, type = "book", issued = new Dictionary<string, object> { ["date-parts"] = new[] { new[] { 2020 } } } })), BibliographyFormat.CslJson).Document;
        string names = emptyNameExpression ? "<names variable=\"author\"><name form=\"short\"/></names>" : string.Empty;
        var style = CslStyle.Parse("<style xmlns=\"http://purl.org/net/xbiblio/csl\" version=\"1.0\" class=\"in-text\"><citation collapse=\"" + mode + "\" disambiguate-add-year-suffix=\"true\" year-suffix-delimiter=\",\" cite-group-delimiter=\", \"><layout delimiter=\"; \">" + names + "<date variable=\"issued\"><date-part name=\"year\"/></date></layout></citation></style>");
        var cite = new CslCitation("one");
        for (int index = 0; index < count; index++) cite.Items.Add(new CslCitationItem("item-" + index) { Prefix = index == prefixIndex ? "see " : null });
        var html = new CslProcessor(data, style, new CslRenderOptions { OutputFormat = CslOutputFormat.Html }).Render(new[] { cite }).Citations.Single().Content;
        var plain = new CslProcessor(data, style, new CslRenderOptions { OutputFormat = CslOutputFormat.PlainText }).Render(new[] { cite }).Citations.Single().Content;
        Assert.Equal(html, plain);
        return plain;
    }
}
