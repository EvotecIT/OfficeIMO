namespace OfficeIMO.Bibliography.Tests;

public sealed class CslSubstitutionConditionTests {
    [Theory]
    [InlineData("author", "\"author\":[{\"literal\":\"Writer\"}]", "<names variable=\"author\"/>", "Writer", false, false)]
    [InlineData("author", "\"author\":[{\"literal\":\"Writer\"}]", "<names variable=\"author\"/>", "Writer", false, true)]
    [InlineData("author", "\"author\":[{\"literal\":\"Writer\"}]", "<names variable=\"author\"/>", "Writer", true, false)]
    [InlineData("author", "\"author\":[{\"literal\":\"Writer\"}]", "<names variable=\"author\"/>", "Writer", true, true)]
    [InlineData("volume", "\"volume\":\"12\"", "<text variable=\"volume\"/>", "12", false, false)]
    [InlineData("volume", "\"volume\":\"12\"", "<text variable=\"volume\"/>", "12", false, true)]
    [InlineData("volume", "\"volume\":\"12\"", "<text variable=\"volume\"/>", "12", true, false)]
    [InlineData("volume", "\"volume\":\"12\"", "<text variable=\"volume\"/>", "12", true, true)]
    public void SubstitutionSuppressesRepeatedRenderingWhileConditionsRetainSourcePresence(string variable, string fields, string rendering, string text, bool bibliography, bool html) {
        string body = "<group delimiter=\"|\"><names variable=\"composer\"><substitute>" + rendering +
            "</substitute></names><choose><if variable=\"" + variable + "\"><text value=\"present\"/></if>" +
            "<else><text value=\"missing\"/></else></choose>" + rendering + "</group>";
        Assert.Equal(text + "|present", Render(fields, body, bibliography, html));
    }

    [Theory]
    [InlineData("12", "numeric", false, false)]
    [InlineData("12", "numeric", false, true)]
    [InlineData("12", "numeric", true, false)]
    [InlineData("12", "numeric", true, true)]
    [InlineData("Special edition", "text", false, false)]
    [InlineData("Special edition", "text", false, true)]
    [InlineData("Special edition", "text", true, false)]
    [InlineData("Special edition", "text", true, true)]
    public void NumericConditionsClassifyTheSourceAfterItWasUsedAsASubstitute(string value, string expected, bool bibliography, bool html) {
        string body = "<group delimiter=\"|\"><names variable=\"composer\"><substitute><text variable=\"volume\"/>" +
            "</substitute></names><choose><if is-numeric=\"volume\"><text value=\"numeric\"/></if>" +
            "<else><text value=\"text\"/></else></choose><text variable=\"volume\"/></group>";
        Assert.Equal(value + "|" + expected, Render("\"volume\":\"" + value + "\"", body, bibliography, html));
    }

    private static string Render(string fields, string body, bool bibliography, bool html) {
        var data = BibliographyDocument.Parse("[{\"id\":\"one\",\"type\":\"book\"," + fields + "}]", BibliographyFormat.CslJson).Document;
        const string header = "<style xmlns=\"http://purl.org/net/xbiblio/csl\" version=\"1.0\" class=\"in-text\">";
        string xml = header + "<citation><layout>" + (bibliography ? "<text variable=\"title\"/>" : body) + "</layout></citation>" +
            (bibliography ? "<bibliography><layout>" + body + "</layout></bibliography>" : "") + "</style>";
        var cite = new CslCitation("cluster"); cite.Items.Add(new CslCitationItem("one"));
        var processor = new CslProcessor(data, CslStyle.Parse(xml), new CslRenderOptions { OutputFormat = html ? CslOutputFormat.Html : CslOutputFormat.PlainText });
        string result = bibliography ? processor.RenderBibliography().Single().Content : processor.Render(new[] { cite }).Citations.Single().Content;
        return bibliography && html ? result.Substring("<div class=\"csl-entry\">".Length, result.Length - "<div class=\"csl-entry\"></div>".Length) : result;
    }
}
