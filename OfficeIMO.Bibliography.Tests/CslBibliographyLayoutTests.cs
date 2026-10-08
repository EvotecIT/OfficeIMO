namespace OfficeIMO.Bibliography.Tests;

public sealed class CslBibliographyLayoutTests {
    private const string Header = "<style xmlns=\"http://purl.org/net/xbiblio/csl\" version=\"1.0\" class=\"in-text\"><citation><layout><text variable=\"title\"/></layout></citation>";
    private static readonly BibliographyDocument Data = BibliographyDocument.Parse("[{\"id\":\"one\",\"type\":\"book\",\"title\":\"Alpha\"},{\"id\":\"two\",\"type\":\"book\",\"title\":\"Beta\"}]", BibliographyFormat.CslJson).Document;

    [Theory]
    [InlineData("flush", CslSecondFieldAlignment.Flush)]
    [InlineData("margin", CslSecondFieldAlignment.Margin)]
    public void SecondFieldAlignmentAndSpacingAreAvailableToTheHost(string alignment, CslSecondFieldAlignment expected) {
        CslStyle style = CslStyle.Parse(Header + "<bibliography second-field-align=\"" + alignment + "\" hanging-indent=\"true\" line-spacing=\"2\" entry-spacing=\"0\"><layout><text variable=\"citation-number\" prefix=\"[\" suffix=\"] \"/><text variable=\"title\" font-style=\"italic\"/></layout></bibliography></style>");
        var processor = new CslProcessor(Data, style, new CslRenderOptions { OutputFormat = CslOutputFormat.Html });
        CslRenderResult result = processor.Render(Array.Empty<CslCitation>(), includeUncitedItems: true);

        Assert.Equal("<div class=\"csl-entry\"><div class=\"csl-left-margin\">[1] </div><div class=\"csl-right-inline\"><i>Alpha</i></div></div>", result.Bibliography[0].Content);
        CslBibliographyLayout layout = Assert.IsType<CslBibliographyLayout>(result.BibliographyLayout);
        Assert.Equal(expected, layout.SecondFieldAlignment);
        Assert.True(layout.HangingIndent);
        Assert.Equal(2, layout.LineSpacing);
        Assert.Equal(0, layout.EntrySpacing);
        Assert.Equal(4, layout.MaximumLeftMarginCharacters);
        Assert.Equal(new[] { "[1] Alpha", "[2] Beta" }, new CslProcessor(Data, style).RenderBibliography().Select(entry => entry.Content));
    }

    [Fact]
    public void DisplayAffixesStayInsideTheFieldAndOutsideItsEmphasis() {
        CslStyle style = CslStyle.Parse(Header + "<bibliography><layout><text variable=\"citation-number\" display=\"left-margin\" font-style=\"italic\" prefix=\"[\" suffix=\"]\"/><text variable=\"title\" display=\"right-inline\"/></layout></bibliography></style>");
        CslRenderResult result = new CslProcessor(Data, style, new CslRenderOptions { OutputFormat = CslOutputFormat.Html }).Render(Array.Empty<CslCitation>(), true);

        Assert.Equal("<div class=\"csl-entry\"><div class=\"csl-left-margin\">[<i>1</i>]</div><div class=\"csl-right-inline\">Alpha</div></div>", result.Bibliography[0].Content);
        Assert.Equal(3, result.BibliographyLayout!.MaximumLeftMarginCharacters);
    }

    [Fact]
    public void LayoutDefaultsExistForAnEmptyBibliographyAndAbsentSectionIsDistinct() {
        var empty = BibliographyDocument.Parse("[]", BibliographyFormat.CslJson).Document;
        CslRenderResult result = new CslProcessor(empty, CslStyle.Parse(Header + "<bibliography><layout><text variable=\"title\"/></layout></bibliography></style>")).Render(Array.Empty<CslCitation>(), true);

        Assert.Empty(result.Bibliography);
        CslBibliographyLayout layout = Assert.IsType<CslBibliographyLayout>(result.BibliographyLayout);
        Assert.Equal(1, layout.LineSpacing);
        Assert.Equal(1, layout.EntrySpacing);
        Assert.False(layout.HangingIndent);
        Assert.Equal(CslSecondFieldAlignment.None, layout.SecondFieldAlignment);
        Assert.Equal(0, layout.MaximumLeftMarginCharacters);
        Assert.Null(new CslProcessor(empty, CslStyle.Parse(Header + "</style>")).Render(Array.Empty<CslCitation>()).BibliographyLayout);
    }

    [Theory]
    [InlineData("line-spacing=\"0\"")]
    [InlineData("entry-spacing=\"-1\"")]
    [InlineData("entry-spacing=\"1.5\"")]
    [InlineData("second-field-align=\"center\"")]
    [InlineData("hanging-indent=\"yes\"")]
    public void InvalidLayoutSettingsFailDuringStyleLoading(string attribute) =>
        Assert.Throws<InvalidDataException>(() => CslStyle.Parse(Header + "<bibliography " + attribute + "><layout><text variable=\"title\"/></layout></bibliography></style>"));

    [Fact]
    public void AlignmentMarkupCountsTowardTheIntermediateBudget() {
        CslStyle style = CslStyle.Parse(Header + "<bibliography second-field-align=\"flush\"><layout><text variable=\"citation-number\"/><text variable=\"title\"/></layout></bibliography></style>");
        var processor = new CslProcessor(Data, style, new CslRenderOptions { OutputFormat = CslOutputFormat.Html, MaximumIntermediateCharacters = 10 });
        Assert.Throws<InvalidDataException>(() => processor.RenderBibliography());
    }
}
