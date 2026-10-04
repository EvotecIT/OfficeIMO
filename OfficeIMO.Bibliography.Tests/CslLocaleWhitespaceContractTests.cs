namespace OfficeIMO.Bibliography.Tests;

public sealed class CslLocaleWhitespaceContractTests {
    [Theory]
    [InlineData(false, false, false)]
    [InlineData(true, false, false)]
    [InlineData(false, true, false)]
    [InlineData(true, true, false)]
    [InlineData(false, true, true)]
    [InlineData(true, true, true)]
    public void IndentationInEmptyLocaleTermsDoesNotBecomeCitationText(bool external, bool variants, bool plural) {
        string content = variants ? "<single>\n  </single><multiple>\n  </multiple>" : "\n  ";
        var result = Render(external, "<term name=\"and\">" + content + "</term>", plural);
        Assert.Empty(result.Content);
        Assert.True(result.IsEmpty);
    }

    [Theory]
    [InlineData(false, "<term name=\"and\"> AD </term>", " AD ")]
    [InlineData(true, "<term name=\"and\"> AD </term>", " AD ")]
    [InlineData(false, "<term name=\"and\" xml:space=\"preserve\"> </term>", " ")]
    [InlineData(true, "<term name=\"and\" xml:space=\"preserve\"> </term>", " ")]
    [InlineData(false, "<term name=\"and\" xml:space=\"preserve\">\n  </term>", "\n  ")]
    [InlineData(true, "<term name=\"and\" xml:space=\"preserve\">\n  </term>", "\n  ")]
    public void MeaningfulTermPaddingAndExplicitPreservationRemainContent(bool external, string term, string expected) {
        var result = Render(external, term, plural: false);
        Assert.Equal(expected, result.Content);
        Assert.False(result.IsEmpty);
    }

    [Theory]
    [InlineData(false, "<term name=\"and\">AD<!--a--> <!--b-->BC</term>", "AD BC")]
    [InlineData(true, "<term name=\"and\">AD<!--a--> <!--b-->BC</term>", "AD BC")]
    [InlineData(false, "<term name=\"and\"> <![CDATA[AD]]> </term>", " AD ")]
    [InlineData(true, "<term name=\"and\"> <![CDATA[AD]]> </term>", " AD ")]
    [InlineData(false, "<term name=\"and\"> <!--a-->AD<!--b--> </term>", " AD ")]
    [InlineData(true, "<term name=\"and\"> <!--a-->AD<!--b--> </term>", " AD ")]
    [InlineData(false, "<term name=\"and\">\u00a0</term>", "&#160;")]
    [InlineData(true, "<term name=\"and\">\u00a0</term>", "&#160;")]
    public void NonemptyTermContentSurvivesXmlNodeBoundaries(bool external, string term, string expected) {
        var result = Render(external, term, plural: false);
        Assert.Equal(expected, result.Content);
        Assert.False(result.IsEmpty);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void IndentationInAnEmptyOrdinalSuffixDoesNotBecomeCitationText(bool external) {
        const string ns = "http://purl.org/net/xbiblio/csl";
        string locale = "<locale xmlns=\"" + ns + "\" xml:lang=\"en-US\"><terms><term name=\"ordinal\">\n  </term></terms></locale>";
        string xml = "<style xmlns=\"" + ns + "\" version=\"1.0\" class=\"in-text\">" + (external ? "" : locale) +
            "<citation><layout><number variable=\"volume\" form=\"ordinal\"/></layout></citation></style>";
        var options = new CslRenderOptions { OutputFormat = CslOutputFormat.Html };
        if (external) options.Locales["en-US"] = locale;
        var data = BibliographyDocument.Parse("[{\"id\":\"one\",\"type\":\"book\",\"volume\":1}]", BibliographyFormat.CslJson).Document;
        var cite = new CslCitation("one"); cite.Items.Add(new CslCitationItem("one"));
        Assert.Equal("1", new CslProcessor(data, CslStyle.Parse(xml), options).Render(new[] { cite }).Citations.Single().Content);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TermCanOverrideAnInheritedWhitespacePreservationScope(bool external) {
        var result = Render(external, "<term name=\"and\" xml:space=\"default\">\n  </term>", plural: false, preserveLocale: true);
        Assert.Empty(result.Content);
        Assert.True(result.IsEmpty);
    }

    private static CslRenderedEntry Render(bool external, string term, bool plural, bool preserveLocale = false) {
        const string ns = "http://purl.org/net/xbiblio/csl";
        string space = preserveLocale ? " xml:space=\"preserve\"" : string.Empty;
        string locale = "<locale xmlns=\"" + ns + "\" xml:lang=\"en-US\"" + space + "><terms>" + term + "</terms></locale>";
        string xml = "<style xmlns=\"" + ns + "\" version=\"1.0\" class=\"in-text\" default-locale=\"en-US\">" +
            (external ? string.Empty : locale) + "<citation><layout><text term=\"and\" plural=\"" + (plural ? "true" : "false") + "\"/></layout></citation></style>";
        var options = new CslRenderOptions { OutputFormat = CslOutputFormat.Html };
        if (external) options.Locales["en-US"] = locale;
        var data = BibliographyDocument.Parse("[{\"id\":\"one\",\"type\":\"book\"}]", BibliographyFormat.CslJson).Document;
        var cite = new CslCitation("one"); cite.Items.Add(new CslCitationItem("one"));
        return new CslProcessor(data, CslStyle.Parse(xml), options).Render(new[] { cite }).Citations.Single();
    }
}
