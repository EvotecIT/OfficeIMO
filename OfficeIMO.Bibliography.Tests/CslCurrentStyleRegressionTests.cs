using System.Text.Json;

namespace OfficeIMO.Bibliography.Tests;

public sealed class CslCurrentStyleRegressionTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AnImplicitNamePrecedesItsContributorLabel(bool html) {
        const string fields = "\"editor\":[{\"family\":\"Chen\",\"given\":\"Wei\"}]";
        string names = "<names variable=\"editor\"><label prefix=\", \"" + (html ? " font-style=\"italic\"" : "") + "/></names>";
        // Names formatting surrounds the implicit name and its separately formatted label.
        if (html) names = names.Replace("<names ", "<names font-weight=\"bold\" ");
        Assert.Equal(html ? "<b>Chen, Wei, <i>editor</i></b>" : "Chen, Wei, editor", Render(fields, names, html, " name-as-sort-order=\"all\""));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AnImplicitInheritedNamePrecedesTheSubstitutedEditorLabel(bool html) {
        const string fields = "\"editor\":[{\"literal\":\"Example Editor\"}]";
        const string names = "<names variable=\"author\"><label form=\"short\" prefix=\", \"/><substitute><names variable=\"editor\"/></substitute></names>";
        Assert.Equal("Example Editor, ed.", Render(fields, names, html));
    }

    [Theory]
    [InlineData("4.2", "version")]
    [InlineData("4.2.1", "version")]
    [InlineData("v4.2", "version")]
    [InlineData("4.2 & 5.3", "versions")]
    [InlineData("4.2-5.3", "versions")]
    public void DottedVersionsAreSingleValuesUnlessAListOrRangeIsPresent(string value, string expected) {
        foreach (bool html in new[] { false, true })
            Assert.Equal(expected, Render("\"version\":" + JsonSerializer.Serialize(value), "<label variable=\"version\"/>", html));
    }

    [Theory]
    [InlineData("number", "ES-22-8")]
    [InlineData("volume", "Report-22-8")]
    [InlineData("issue", "Special-22-8")]
    [InlineData("edition", "Revised-22-8")]
    public void TextIdentifiersWithNonnumericContentRetainTheirHyphens(string variable, string value) {
        foreach (bool html in new[] { false, true })
            Assert.Equal(value, Render("\"" + variable + "\":" + JsonSerializer.Serialize(value), "<text variable=\"" + variable + "\"/>", html));
    }

    private static string Render(string fields, string body, bool html, string styleOptions = "") {
        var data = BibliographyDocument.Parse("[{\"id\":\"one\",\"type\":\"chapter\"," + fields + "}]", BibliographyFormat.CslJson).Document;
        var style = CslStyle.Parse("<style xmlns=\"http://purl.org/net/xbiblio/csl\" version=\"1.0\" class=\"in-text\"" + styleOptions + "><citation><layout>" + body + "</layout></citation></style>");
        var cite = new CslCitation("one"); cite.Items.Add(new CslCitationItem("one"));
        return new CslProcessor(data, style, new CslRenderOptions { OutputFormat = html ? CslOutputFormat.Html : CslOutputFormat.PlainText }).Render(new[] { cite }).Citations.Single().Content;
    }
}
