namespace OfficeIMO.Bibliography.Tests;

public sealed class CslDisplayProjectionTests {
    private const string Header = "<style xmlns=\"http://purl.org/net/xbiblio/csl\" version=\"1.0\" class=\"in-text\"><citation><layout><text variable=\"title\"/></layout></citation>";
    private static readonly BibliographyDocument Data = BibliographyDocument.Parse("[{\"id\":\"one\",\"type\":\"book\",\"title\":\"Alpha\"}]", BibliographyFormat.CslJson).Document;

    [Theory]
    [InlineData("flush", CslOutputFormat.Html)]
    [InlineData("margin", CslOutputFormat.Html)]
    [InlineData("flush", CslOutputFormat.PlainText)]
    [InlineData("margin", CslOutputFormat.PlainText)]
    public void LayoutAffixesBelongToAlignedFieldsAndTheWidthHint(string alignment, CslOutputFormat format) {
        string body = "<layout prefix=\"(\" suffix=\").\"><text variable=\"citation-number\" prefix=\"[\" suffix=\"]\"/><text variable=\"title\" font-style=\"italic\"/></layout>";
        CslRenderResult result = Render(body, "second-field-align=\"" + alignment + "\"", format);
        string expected = format == CslOutputFormat.Html ? "<div class=\"csl-entry\"><div class=\"csl-left-margin\">([1]</div><div class=\"csl-right-inline\"><i>Alpha</i>).</div></div>" : "([1]Alpha).";
        Assert.Equal(expected, result.Bibliography.Single().Content);
        Assert.Equal(4, result.BibliographyLayout!.MaximumLeftMarginCharacters);
    }

    [Fact]
    public void LayoutFormattingAndQuotesApplyBeforeSplittingAlignedFields() {
        const string body = "<layout prefix=\"pre.(\" suffix=\").\" strip-periods=\"true\" text-case=\"uppercase\" font-style=\"italic\" quotes=\"true\"><text value=\"ID.\"/><text variable=\"title\"/></layout>";
        CslRenderResult result = Render(body, "second-field-align=\"flush\"");
        Assert.Equal("<div class=\"csl-entry\"><div class=\"csl-left-margin\">“<i>PRE(ID</i></div><div class=\"csl-right-inline\"><i>ALPHA)</i>”</div></div>", result.Bibliography.Single().Content);
        Assert.Equal(7, result.BibliographyLayout!.MaximumLeftMarginCharacters);
        Assert.Equal("“PRE(IDALPHA)”", Render(body, "second-field-align=\"flush\"", CslOutputFormat.PlainText).Bibliography.Single().Content);
    }

    [Fact]
    public void LayoutEmphasisAndInputFlipsSurviveProjection() {
        var data = BibliographyDocument.Parse("[{\"id\":\"one\",\"type\":\"book\",\"title\":\"<i>Alpha</i>\"}]", BibliographyFormat.CslJson).Document;
        const string body = "<layout font-style=\"italic\" suffix=\".\"><text variable=\"citation-number\" prefix=\"[\" suffix=\"]\"/><text variable=\"title\"/></layout>";
        Assert.Equal("<div class=\"csl-entry\"><div class=\"csl-left-margin\"><i>[1]</i></div><div class=\"csl-right-inline\"><i><span style=\"font-style:normal;\">Alpha</span>.</i></div></div>", Render(body, "second-field-align=\"flush\"", data: data).Bibliography.Single().Content);
    }

    [Theory]
    [InlineData(CslOutputFormat.Html)]
    [InlineData(CslOutputFormat.PlainText)]
    public void ExplicitDisplayBlocksOwnLayoutAffixesAndFinalWidth(CslOutputFormat format) {
        const string body = "<layout prefix=\"(\" suffix=\").\"><text variable=\"citation-number\" prefix=\"[\" suffix=\"]\" display=\"left-margin\"/><text variable=\"title\" font-style=\"italic\" display=\"right-inline\"/></layout>";
        CslRenderResult result = Render(body, format: format);
        string expected = format == CslOutputFormat.Html ? "<div class=\"csl-entry\"><div class=\"csl-left-margin\">([1]</div><div class=\"csl-right-inline\"><i>Alpha</i>).</div></div>" : "([1]Alpha).";
        Assert.Equal(expected, result.Bibliography.Single().Content);
        Assert.Equal(4, result.BibliographyLayout!.MaximumLeftMarginCharacters);
    }

    [Fact]
    public void WidthHintCountsFinalTextAfterParentTransformAndPunctuationCleanup() {
        const string body = "<layout strip-periods=\"true\"><text value=\"ID...\" display=\"left-margin\"/><text variable=\"title\" display=\"right-inline\"/></layout>";
        CslRenderResult result = Render(body);
        Assert.Equal("<div class=\"csl-entry\"><div class=\"csl-left-margin\">ID</div><div class=\"csl-right-inline\">Alpha</div></div>", result.Bibliography.Single().Content);
        Assert.Equal(2, result.BibliographyLayout!.MaximumLeftMarginCharacters);
    }

    [Fact]
    public void MultipleDisplayBlocksRetainLayoutFormattingAndOnePairOfQuotes() {
        const string body = "<layout prefix=\"(\" suffix=\").\" font-style=\"italic\" quotes=\"true\"><text value=\"Heading\" display=\"block\"/><text variable=\"title\" display=\"block\"/><text value=\"End\" display=\"indent\"/></layout>";
        CslRenderResult result = Render(body);
        Assert.Equal("<div class=\"csl-entry\"><div class=\"csl-block\">“<i>(Heading</i></div><div class=\"csl-block\"><i>Alpha</i></div><div class=\"csl-indent\"><i>End).</i>”</div></div>", result.Bibliography.Single().Content);
        Assert.Equal("“(HeadingAlphaEnd).”", Render(body, format: CslOutputFormat.PlainText).Bibliography.Single().Content);
        Assert.Equal(0, result.BibliographyLayout!.MaximumLeftMarginCharacters);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MixedInlineAndDisplayFieldsKeepTheirOriginalPositions(bool grouped) {
        const string fields = "<text value=\"Start\"/><text value=\"Block\" display=\"indent\"/><text value=\"End\"/>";
        string body = "<layout prefix=\"(\" suffix=\").\" font-style=\"italic\" quotes=\"true\">" + (grouped ? "<group>" + fields + "</group>" : fields) + "</layout>";
        CslRenderResult result = Render(body);
        Assert.Equal("<div class=\"csl-entry\">“<i>(Start</i><div class=\"csl-indent\"><i>Block</i></div><i>End).</i>”</div>", result.Bibliography.Single().Content);
        Assert.Equal("“(StartBlockEnd).”", Render(body, format: CslOutputFormat.PlainText).Bibliography.Single().Content);
    }

    [Fact]
    public void SingleFieldKeepsLayoutAffixesWithoutCreatingAlignmentColumns() {
        const string body = "<layout prefix=\"(\" suffix=\").\"><text variable=\"citation-number\" prefix=\"[\" suffix=\"]\"/><text variable=\"author\"/></layout>";
        CslRenderResult result = Render(body, "second-field-align=\"flush\"");
        Assert.Equal("<div class=\"csl-entry\">([1]).</div>", result.Bibliography.Single().Content);
        Assert.Equal(0, result.BibliographyLayout!.MaximumLeftMarginCharacters);
    }

    [Fact]
    public void LongestWidthComesFromFinalVisibleColumnsAcrossEntries() {
        var data = BibliographyDocument.Parse("[{\"id\":\"one\",\"type\":\"book\",\"title\":\"ID...\"},{\"id\":\"two\",\"type\":\"book\",\"title\":\"LONG...\"}]", BibliographyFormat.CslJson).Document;
        const string body = "<layout prefix=\"(\" strip-periods=\"true\"><text variable=\"title\" display=\"left-margin\"/><text value=\"Ref\" display=\"right-inline\"/></layout>";
        CslRenderResult result = Render(body, data: data);
        Assert.Equal(5, result.BibliographyLayout!.MaximumLeftMarginCharacters);
        Assert.Contains("<div class=\"csl-left-margin\">(ID</div>", result.Bibliography[0].Content);
        Assert.Contains("<div class=\"csl-left-margin\">(LONG</div>", result.Bibliography[1].Content);
        Assert.Equal(5, Render(body, format: CslOutputFormat.PlainText, data: data).BibliographyLayout!.MaximumLeftMarginCharacters);
    }

    [Fact]
    public void IdentifierLinksKeepTheirTargetAndFormattingInsideTheBodyColumn() {
        var data = BibliographyDocument.Parse("[{\"id\":\"one\",\"type\":\"book\",\"title\":\"Alpha\",\"DOI\":\"10.1234/example\"}]", BibliographyFormat.CslJson).Document;
        const string body = "<layout suffix=\".\" font-style=\"italic\"><text variable=\"citation-number\"/><group delimiter=\"; \"><text variable=\"title\"/><text variable=\"DOI\"/></group></layout>";
        CslRenderResult result = Render(body, "second-field-align=\"flush\"", data: data);
        Assert.Equal("<div class=\"csl-entry\"><div class=\"csl-left-margin\"><i>1</i></div><div class=\"csl-right-inline\"><i>Alpha; <a href=\"https://doi.org/10.1234/example\">10.1234/example</a>.</i></div></div>", result.Bibliography.Single().Content);
    }

    [Fact]
    public void ProjectedAncestorFormattingCountsTowardTheIntermediateBudget() {
        const string body = "<layout font-style=\"italic\"><text variable=\"citation-number\"/><text variable=\"title\"/></layout>";
        CslStyle style = CslStyle.Parse(Header + "<bibliography second-field-align=\"flush\">" + body + "</bibliography></style>");
        var processor = new CslProcessor(Data, style, new CslRenderOptions { OutputFormat = CslOutputFormat.Html, MaximumIntermediateCharacters = 90 });
        InvalidDataException error = Assert.Throws<InvalidDataException>(() => processor.RenderBibliography());
        Assert.Contains("MaximumIntermediateCharacters", error.Message);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void GroupOwnedAffixesStayWithItsDisplayFieldsAndOutsideItsEmphasis(bool surroundingInlineFields) {
        string body = "<group prefix=\"(\" suffix=\")\" font-style=\"italic\"><text value=\"ID\" display=\"left-margin\"/><text variable=\"title\" display=\"right-inline\"/></group>";
        if (surroundingInlineFields) body = "<text value=\"Before\"/>" + body + "<text value=\"After\"/>";
        CslRenderResult result = Render("<layout>" + body + "</layout>");
        string expected = "<div class=\"csl-left-margin\">(<i>ID</i></div><div class=\"csl-right-inline\"><i>Alpha</i>)</div>";
        if (surroundingInlineFields) expected = "Before" + expected + "After";
        Assert.Equal("<div class=\"csl-entry\">" + expected + "</div>", result.Bibliography.Single().Content);
        Assert.Equal(3, result.BibliographyLayout!.MaximumLeftMarginCharacters);
    }

    [Fact]
    public void MacroWrapperQuotesAndAffixesRetainTheirOwnedColumns() {
        const string macro = "<macro name=\"fields\"><text value=\"ID\" display=\"left-margin\"/><text variable=\"title\" display=\"right-inline\"/></macro>";
        var style = CslStyle.Parse(Header + macro + "<bibliography><layout><text macro=\"fields\" prefix=\"(\" suffix=\")\" quotes=\"true\" font-style=\"italic\"/></layout></bibliography></style>");
        CslRenderResult result = new CslProcessor(Data, style, new CslRenderOptions { OutputFormat = CslOutputFormat.Html }).Render(Array.Empty<CslCitation>(), true);
        Assert.Equal("<div class=\"csl-entry\"><div class=\"csl-left-margin\">(“<i>ID</i></div><div class=\"csl-right-inline\"><i>Alpha</i>”)</div></div>", result.Bibliography.Single().Content);
        Assert.Equal(4, result.BibliographyLayout!.MaximumLeftMarginCharacters);
        Assert.Equal("(“IDAlpha”)", new CslProcessor(Data, style).RenderBibliography().Single().Content);
    }

    [Fact]
    public void SeparateDisplayWrappersDoNotExchangeTheirAffixes() {
        const string group = "<group prefix=\"(\" suffix=\")\"><text value=\"ID\" display=\"left-margin\"/><text variable=\"title\" display=\"right-inline\"/></group>";
        CslRenderResult result = Render("<layout>" + group + group.Replace("ID", "Key") + "</layout>");
        Assert.Equal("<div class=\"csl-entry\"><div class=\"csl-left-margin\">(ID</div><div class=\"csl-right-inline\">Alpha)</div><div class=\"csl-left-margin\">(Key</div><div class=\"csl-right-inline\">Alpha)</div></div>", result.Bibliography.Single().Content);
        Assert.Equal(4, result.BibliographyLayout!.MaximumLeftMarginCharacters);
    }

    private static CslRenderResult Render(string layout, string settings = "", CslOutputFormat format = CslOutputFormat.Html, BibliographyDocument? data = null) {
        CslStyle style = CslStyle.Parse(Header + "<bibliography " + settings + ">" + layout + "</bibliography></style>");
        return new CslProcessor(data ?? Data, style, new CslRenderOptions { OutputFormat = format }).Render(Array.Empty<CslCitation>(), true);
    }
}
