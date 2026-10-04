namespace OfficeIMO.Bibliography.Tests;

public sealed class CslReviewRegressionTests {
    private const string Header = "<style xmlns=\"http://purl.org/net/xbiblio/csl\" version=\"1.0\" class=\"in-text\">";
    private const string IdLayout = "<layout><text variable=\"id\"/></layout>";

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MacroDependencyDepthIsIndependentOfDeclarationOrder(bool reverse) {
        string[] macros = Enumerable.Range(0, 12).Select(index => "<macro name=\"m" + index + "\">" +
            (index == 0 ? "<text variable=\"title\"/>" : "<text macro=\"m" + (index - 1) + "\"/>") + "</macro>").ToArray();
        if (reverse) Array.Reverse(macros);
        string xml = Header + string.Concat(macros) + "<citation><layout><text macro=\"m11\"/></layout></citation></style>";
        Assert.Throws<InvalidDataException>(() => CslStyle.Parse(xml, new CslStyleLoadOptions { MaximumNestingDepth = 4 }));
    }

    [Fact]
    public void SharedMacroDependenciesRetainExplicitYearSuffixDiscovery() {
        CslStyle style = CslStyle.Parse(Header + """
            <macro name="leaf"><text variable="year-suffix"/></macro>
            <macro name="left"><text macro="leaf"/></macro>
            <macro name="right"><text macro="leaf"/></macro>
            <citation disambiguate-add-year-suffix="true"><layout><date variable="issued"><date-part name="year"/></date><text macro="right"/><text macro="left"/></layout></citation>
            </style>
            """);
        BibliographyDocument data = Data("[{\"id\":\"a\",\"type\":\"book\",\"issued\":{\"date-parts\":[[2024]]}},{\"id\":\"b\",\"type\":\"book\",\"issued\":{\"date-parts\":[[2024]]}}]");
        var citation = new CslCitation("cite"); citation.Items.Add(new CslCitationItem("a")); citation.Items.Add(new CslCitationItem("b"));
        Assert.Equal("2024aa2024bb", new CslProcessor(data, style).Render(new[] { citation }).Citations.Single().Content);
    }

    [Theory]
    [InlineData("font-style=\"italic; background:url(https://attacker.example/x)\"")]
    [InlineData("font-weight=\"bold; position:fixed\"")]
    [InlineData("display=\"block; position:fixed\"")]
    public void ExternalLocaleDateTemplatesRejectInvalidPresentation(string attribute) {
        CslStyle style = CslStyle.Parse(Header + "<citation><layout><date variable=\"issued\" form=\"text\"/></layout></citation></style>");
        var options = new CslRenderOptions { OutputFormat = CslOutputFormat.Html };
        options.Locales.Add("en-US", "<locale xmlns=\"http://purl.org/net/xbiblio/csl\"><date form=\"text\"><date-part name=\"year\" " + attribute + "/></date></locale>");
        var citation = new CslCitation("cite"); citation.Items.Add(new CslCitationItem("a"));
        var processor = new CslProcessor(Data("[{\"id\":\"a\",\"type\":\"book\",\"issued\":{\"date-parts\":[[2024]]}}]"), style, options);
        Assert.Throws<InvalidDataException>(() => processor.Render(new[] { citation }));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ScalarSortKeysEnforceBothIntermediateRepresentations(bool escaped) {
        string title = escaped ? string.Concat(Enumerable.Repeat("&amp;", 20)) : new string('x', 100);
        CslStyle style = CslStyle.Parse(Header + "<citation>" + IdLayout + "</citation><bibliography><sort><key variable=\"title\"/></sort>" + IdLayout + "</bibliography></style>");
        var processor = new CslProcessor(Data("[{\"id\":\"a\",\"type\":\"book\",\"title\":\"" + title + "\"}]"), style,
            new CslRenderOptions { MaximumIntermediateCharacters = 20 });
        InvalidDataException error = Assert.Throws<InvalidDataException>(() => processor.RenderBibliography());
        Assert.Contains("MaximumIntermediateCharacters", error.Message);
    }

    [Theory]
    [InlineData(2010)]
    [InlineData(2020)]
    public void DateContainingSortMacrosRetainNonDateComponents(int zuluYear) {
        CslStyle style = CslStyle.Parse(Header + "<macro name=\"title\"><text variable=\"title\"/></macro><macro name=\"sort\"><group><text macro=\"title\"/><date variable=\"issued\"><date-part name=\"year\"/></date></group></macro><citation>" + IdLayout + "</citation><bibliography><sort><key macro=\"sort\"/></sort>" + IdLayout + "</bibliography></style>");
        BibliographyDocument data = Data("[{\"id\":\"z\",\"type\":\"book\",\"title\":\"Zulu\",\"issued\":{\"date-parts\":[[" + zuluYear + "]]}},{\"id\":\"a\",\"type\":\"book\",\"title\":\"Alpha\",\"issued\":{\"date-parts\":[[2020]]}}]");
        Assert.Equal(new[] { "a", "z" }, new CslProcessor(data, style).RenderBibliography().Select(entry => entry.Key));
    }

    [Theory]
    [InlineData("de", null, "dialect")]
    [InlineData("en-US", "de", "dialect")]
    [InlineData("de-AT", null, "generic")]
    public void LanguageOnlySelectionsUseThePrimaryDialectOverride(string defaultLocale, string? requestedLocale, string expected) {
        CslStyle style = CslStyle.Parse(Header.Replace("class=\"in-text\"", "class=\"in-text\" default-locale=\"" + defaultLocale + "\"") +
            "<locale xml:lang=\"de-DE\"><terms><term name=\"and\">dialect</term></terms></locale><locale xml:lang=\"de\"><terms><term name=\"and\">generic</term></terms></locale><citation><layout><text term=\"and\"/></layout></citation></style>");
        var citation = new CslCitation("cite"); citation.Items.Add(new CslCitationItem("a"));
        Assert.Equal(expected, new CslProcessor(Data("[{\"id\":\"a\",\"type\":\"book\"}]"), style, new CslRenderOptions { Locale = requestedLocale }).Render(new[] { citation }).Citations.Single().Content);
    }

    private static BibliographyDocument Data(string json) => BibliographyDocument.Parse(json, BibliographyFormat.CslJson).Document;

    [Fact]
    public void LanguageOnlySelectionsPrioritizeCallerPrimaryDialectData() {
        CslStyle style = CslStyle.Parse(Header + "<citation><layout><text term=\"and\"/></layout></citation></style>");
        var options = new CslRenderOptions { Locale = "de" };
        options.Locales.Add("de-DE", "<locale xmlns=\"http://purl.org/net/xbiblio/csl\"><terms><term name=\"and\">primary</term></terms></locale>");
        options.Locales.Add("de", "<locale xmlns=\"http://purl.org/net/xbiblio/csl\"><terms><term name=\"and\">generic</term></terms></locale>");
        var citation = new CslCitation("cite"); citation.Items.Add(new CslCitationItem("a"));
        Assert.Equal("primary", new CslProcessor(Data("[{\"id\":\"a\",\"type\":\"book\"}]"), style, options).Render(new[] { citation }).Citations.Single().Content);
    }

    [Fact]
    public void ValidExternalDatePresentationRemainsAvailable() {
        CslStyle style = CslStyle.Parse(Header + "<citation><layout><date variable=\"issued\" form=\"text\"/></layout></citation></style>");
        var options = new CslRenderOptions { OutputFormat = CslOutputFormat.Html };
        options.Locales.Add("en-US", "<locale xmlns=\"http://purl.org/net/xbiblio/csl\"><date form=\"text\"><date-part name=\"year\" font-style=\"italic\"/></date></locale>");
        var citation = new CslCitation("cite"); citation.Items.Add(new CslCitationItem("a"));
        Assert.Equal("<i>2024</i>", new CslProcessor(Data("[{\"id\":\"a\",\"type\":\"book\",\"issued\":{\"date-parts\":[[2024]]}}]"), style, options).Render(new[] { citation }).Citations.Single().Content);
    }
}
