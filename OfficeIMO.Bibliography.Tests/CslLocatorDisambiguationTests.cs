namespace OfficeIMO.Bibliography.Tests;

public sealed class CslLocatorDisambiguationTests {
    private const string Names = "<names variable=\"author\"><name form=\"short\"/></names>";
    private const string Detail = "<group delimiter=\", \">" + Names + "<choose><if disambiguate=\"true\"><text variable=\"title\"/></if></choose></group>";
    private const string Header = "<style xmlns=\"http://purl.org/net/xbiblio/csl\" version=\"1.0\" class=\"in-text\">";

    [Theory]
    [InlineData("page", false)]
    [InlineData("line", false)]
    [InlineData("page", true)]
    [InlineData("line", true)]
    public void LocatorDependentMacroCallsRetainTheDetailNeededToIdentifyTheWork(string type, bool html) {
        string xml = Header + "<macro name=\"author\">" + Detail + "</macro><citation><layout prefix=\"(\" suffix=\")\"><choose><if locator=\"page line\" match=\"any\"><group delimiter=\" \"><text macro=\"author\"/><text variable=\"locator\"/></group></if><else><text macro=\"author\"/></else></choose></layout></citation></style>";
        var result = Render(xml, html, type, "31", "32");
        Assert.Equal(new[] { "(Doe, Alpha)", "(Doe, Beta)", "(Doe, Alpha 31)", "(Doe, Beta 32)" }, result.Citations.Select(entry => entry.Content));
    }

    [Theory]
    [InlineData("31", "32", "(Doe, Alpha 31)", "(Doe, Beta 32)")]
    [InlineData("front", "back", "(Doe, Alpha front)", "(Doe, Beta back)")]
    public void NumericAndTextualLocatorBranchesAreQualifiedIndependently(string first, string second, string expectedFirst, string expectedSecond) {
        string xml = Header + "<macro name=\"numeric\">" + Detail + "</macro><macro name=\"textual\">" + Detail + "</macro><citation><layout prefix=\"(\" suffix=\")\"><group delimiter=\" \"><choose><if is-numeric=\"locator\"><text macro=\"numeric\"/></if><else><text macro=\"textual\"/></else></choose><text variable=\"locator\"/></group></layout></citation></style>";
        var result = Render(xml, false, "page", first, second);
        Assert.Equal(expectedFirst, result.Citations[2].Content);
        Assert.Equal(expectedSecond, result.Citations[3].Content);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AUniqueLocatorFormDoesNotRepeatConditionalTitleDetail(bool html) {
        string xml = Header + "<macro name=\"author\">" + Detail + "</macro><citation><layout prefix=\"(\" suffix=\")\"><choose><if locator=\"section\"><group delimiter=\" \"><text macro=\"author\"/><text variable=\"title\"/></group></if><else><text macro=\"author\"/></else></choose></layout></citation></style>";
        var result = Render(xml, html, "section", "31", "32");
        Assert.Equal(new[] { "(Doe, Alpha)", "(Doe, Beta)", "(Doe Alpha)", "(Doe Beta)" }, result.Citations.Select(entry => entry.Content));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData(" ")]
    public void ALocatorRequiresAUsableTypeBeforeRendering(string? type) {
        string xml = Header + "<citation><layout><text variable=\"title\"/></layout></citation></style>";
        Assert.Throws<ArgumentException>(() => Render(xml, false, type!, "31", "32"));
    }

    private static CslRenderResult Render(string xml, bool html, string type, string firstLocator, string secondLocator) {
        var data = BibliographyDocument.Parse("[{\"id\":\"a\",\"type\":\"book\",\"author\":[{\"family\":\"Doe\"}],\"title\":\"Alpha\"},{\"id\":\"b\",\"type\":\"book\",\"author\":[{\"family\":\"Doe\"}],\"title\":\"Beta\"}]", BibliographyFormat.CslJson).Document;
        var cites = new[] { new CslCitation("first-a"), new CslCitation("first-b"), new CslCitation("later-a"), new CslCitation("later-b") };
        cites[0].Items.Add(new CslCitationItem("a")); cites[1].Items.Add(new CslCitationItem("b"));
        cites[2].Items.Add(new CslCitationItem("a") { LocatorType = type, Locator = firstLocator });
        cites[3].Items.Add(new CslCitationItem("b") { LocatorType = type, Locator = secondLocator });
        return new CslProcessor(data, CslStyle.Parse(xml), new CslRenderOptions { OutputFormat = html ? CslOutputFormat.Html : CslOutputFormat.PlainText }).Render(cites);
    }
}
