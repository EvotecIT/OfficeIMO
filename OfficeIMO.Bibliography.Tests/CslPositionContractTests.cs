namespace OfficeIMO.Bibliography.Tests;

public sealed class CslPositionContractTests {
    private const string Positions = "<choose><if position=\"ibid-with-locator\"><text value=\"changed\"/></if><else-if position=\"ibid\"><text value=\"repeat\"/></else-if><else-if position=\"first\"><text value=\"first\"/></else-if><else><text value=\"later\"/></else></choose>";
    private static readonly BibliographyDocument Data = BibliographyDocument.Parse("[{\"id\":\"a\",\"type\":\"book\",\"title\":\"alpha\"},{\"id\":\"b\",\"type\":\"book\",\"title\":\"beta\"}]", BibliographyFormat.CslJson).Document;
    private static CslStyle Style(string body = Positions, string citationOptions = "", string styleClass = "note", string locale = "") => CslStyle.Parse(
        "<style xmlns=\"http://purl.org/net/xbiblio/csl\" version=\"1.0\" class=\"" + styleClass + "\">" + locale +
        "<citation " + citationOptions + "><layout delimiter=\"; \">" + body + "</layout></citation></style>");
    private static CslCitation Cite(string id, int note, params string[] keys) {
        var cite = new CslCitation(id) { NoteIndex = note };
        foreach (string key in keys) cite.Items.Add(new CslCitationItem(key));
        return cite;
    }
    private static string[] Render(CslStyle style, params CslCitation[] cites) => new CslProcessor(Data, style).Render(cites).Citations.Select(entry => entry.Content).ToArray();

    [Theory]
    [InlineData(null, "", "repeat")]
    [InlineData("", null, "repeat")]
    [InlineData("", "", "repeat")]
    [InlineData(null, null, "repeat")]
    [InlineData(null, "12", "changed")]
    [InlineData("12", "12", "repeat")]
    [InlineData("12", "13", "changed")]
    [InlineData("12", null, "later")]
    [InlineData("12", "", "later")]
    public void LocatorPresenceControlsRepeatPosition(string? prior, string? current, string expected) {
        CslCitation first = Cite("one", 0, "a"), second = Cite("two", 0, "a");
        first.Items[0].Locator = prior; second.Items[0].Locator = current;
        Assert.Equal(new[] { "first", expected }, Render(Style(styleClass: "in-text"), first, second));
    }

    [Fact]
    public void AbsentLocatorLabelsDoNotChangeRepeatPosition() {
        CslCitation first = Cite("one", 0, "a"), second = Cite("two", 0, "a");
        second.Items[0].LocatorType = "volume";
        Assert.Equal("repeat", Render(Style(styleClass: "in-text"), first, second)[1]);
        first.Items[0].Locator = second.Items[0].Locator = "2";
        Assert.Equal("changed", Render(Style(styleClass: "in-text"), first, second)[1]);
    }

    [Theory]
    [InlineData(0, 1, "First")]
    [InlineData(1, 0, "first")]
    public void BodyAndNotePositionsHaveSeparateHistories(int firstNote, int nextNote, string expected) =>
        Assert.Equal(expected, Render(Style(), Cite("one", firstNote, "a"), Cite("two", nextNote, "a"))[1]);

    [Fact]
    public void BodyCitationsContinueAcrossInterveningNotes() => Assert.Equal(new[] { "first", "First", "repeat" },
        Render(Style(), Cite("body-one", 0, "a"), Cite("note", 1, "b"), Cite("body-two", 0, "a")));

    [Fact]
    public void MultipleWorksInThePreviousNotePreventAnAmbiguousBackreference() => Assert.Equal(new[] { "First", "first", "Later", "Repeat" },
        Render(Style(), Cite("one-a", 1, "a"), Cite("one-b", 1, "b"), Cite("two", 2, "b"), Cite("three", 3, "b")));

    [Fact]
    public void RepeatedCitesToOneWorkInThePreviousNoteAreUnambiguous() => Assert.Equal(new[] { "First", "repeat", "Repeat" },
        Render(Style(), Cite("one", 1, "a"), Cite("two", 1, "a"), Cite("three", 2, "a")));

    [Fact]
    public void MissingNotesDoNotMakeDistantNotesIbid() => Assert.Equal(new[] { "First", "Later" },
        Render(Style(), Cite("one", 1, "a"), Cite("three", 3, "a")));

    [Fact]
    public void MultiItemClustersAndStyleSortingUseTheirRenderedPredecessor() {
        CslCitation one = Cite("one", 1, "a", "b"), two = Cite("two", 2, "b", "b", "a");
        Assert.Equal(new[] { "First; first", "Later; repeat; later" }, Render(Style(), one, two));
        CslStyle sorted = CslStyle.Parse("<style xmlns=\"http://purl.org/net/xbiblio/csl\" version=\"1.0\" class=\"in-text\"><citation><sort><key variable=\"title\"/></sort><layout delimiter=\"; \">" + Positions + "</layout></citation></style>");
        Assert.Equal("first; repeat; first", Render(sorted, Cite("sorted", 0, "a", "b", "a"))[0]);
    }

    [Theory]
    [InlineData(6, "near")]
    [InlineData(7, "far")]
    public void BodyCitationsDoNotEraseTheLastNoteDistance(int nextNote, string expected) {
        const string body = "<choose><if position=\"near-note\"><text value=\"near\"/></if><else><text value=\"far\"/></else></choose>";
        string[] result = Render(Style(body), Cite("note-one", 1, "a"), Cite("body", 0, "a"), Cite("note-next", nextNote, "a"));
        Assert.Equal(expected, result[2].ToLowerInvariant());
        Assert.Equal("far", result[1]);
    }

    [Fact]
    public void FirstNoteReferencesAreEmptyInTheFirstNoteAndBody() {
        CslStyle style = Style("<text variable=\"title\"/><text variable=\"first-reference-note-number\" prefix=\"@\"/>");
        Assert.Equal(new[] { "alpha", "Alpha", "alpha", "alpha", "Alpha@2" }, Render(style,
            Cite("body-first", 0, "a"), Cite("note-first", 2, "a"), Cite("note-again", 2, "a"), Cite("body-again", 0, "a"), Cite("note-next", 3, "a")));
        Assert.Equal(new[] { "alpha", "alpha" }, Render(Style("<text variable=\"title\"/><text variable=\"first-reference-note-number\" prefix=\"@\"/>", styleClass: "in-text"),
            Cite("first", 1, "a"), Cite("next", 2, "a")));
    }

    [Fact]
    public void ReplacingTheDocumentRecalculatesPositionAndFirstNoteReferences() {
        var processor = new CslProcessor(Data, Style("<text variable=\"title\"/><text variable=\"first-reference-note-number\" prefix=\"@\"/>"));
        CslCitation one = Cite("one", 1, "a"), two = Cite("two", 2, "a");
        Assert.Equal("Alpha@1", processor.Render(new[] { one, two }).Citations[1].Content);
        Assert.Equal("Alpha", processor.Render(new[] { two }).Citations[0].Content);
        Assert.Equal("Alpha@1", processor.Render(new[] { one, two }).Citations[1].Content);
    }

    [Theory]
    [InlineData(CslOutputFormat.PlainText, "“Ibid.”")]
    [InlineData(CslOutputFormat.Html, "“<i>Ibid.</i>”")]
    public void NoteOpeningCapitalizationPreservesFormatting(CslOutputFormat format, string expected) {
        var processor = new CslProcessor(Data, Style("<text term=\"ibid\" font-style=\"italic\" quotes=\"true\"/>"), new CslRenderOptions { OutputFormat = format });
        Assert.Equal(expected, processor.Render(new[] { Cite("note", 1, "a") }).Citations[0].Content);
    }

    [Theory]
    [InlineData("<span class=\"nocase\">ibid</span> lower", "ibid lower")]
    [InlineData("<sup>i</sup> lower", "i lower")]
    [InlineData("4 lower", "4 lower")]
    [InlineData("\U00010428 lower", "\U00010400 lower")]
    [InlineData("<span class=\"nocase\">\U00010428</span> lower", "\U00010428 lower")]
    [InlineData("<sup>²</sup> lower", "² lower")]
    [InlineData("² lower", "² lower")]
    [InlineData("ⅰ lower", "ⅰ lower")]
    public void UnicodeAndProtectedNoteOpeningsDoNotCapitalizeLaterWords(string title, string expected) {
        BibliographyDocument data = BibliographyDocument.Parse("[{\"id\":\"a\",\"type\":\"book\",\"title\":" + System.Text.Json.JsonSerializer.Serialize(title) + "}]", BibliographyFormat.CslJson).Document;
        Assert.Equal(expected, new CslProcessor(data, Style("<text variable=\"title\"/>")).Render(new[] { Cite("one", 1, "a") }).Citations[0].Content);
    }

    [Fact]
    public void LiteralPrefixesAndLaterCitationsRemainAsWritten() {
        CslCitation one = Cite("one", 1, "a"), two = Cite("two", 1, "a");
        one.Items[0].Prefix = "see ";
        Assert.Equal(new[] { "see alpha", "alpha" }, Render(Style("<text variable=\"title\"/>"), one, two));
    }

    [Fact]
    public void EmptyClustersDoNotInterruptTheCitationSequenceOrConsumeTheNoteOpening() =>
        Assert.Equal(new[] { "First", "", "Repeat" }, Render(Style(), Cite("one", 1, "a"), Cite("empty", 2), Cite("two", 2, "a")));

    [Fact]
    public void NoteOpeningUsesTheOutputLocalesUnicodeCasing() {
        var processor = new CslProcessor(Data, Style("<text value=\"ibid\"/>"), new CslRenderOptions { Locale = "tr-TR" });
        Assert.Equal("İbid", processor.Render(new[] { Cite("one", 1, "a") }).Citations[0].Content);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FirstNoteNumbersAreAvailableToVariableAndMacroSortKeys(bool macro) {
        string key = macro ? "macro=\"first-note\"" : "variable=\"first-reference-note-number\"";
        CslStyle style = CslStyle.Parse("<style xmlns=\"http://purl.org/net/xbiblio/csl\" version=\"1.0\" class=\"note\"><macro name=\"first-note\"><text variable=\"first-reference-note-number\"/></macro><citation><sort><key " + key + "/></sort><layout delimiter=\"; \"><text variable=\"title\"/><text variable=\"first-reference-note-number\" prefix=\"@\"/></layout></citation></style>");
        Assert.Equal("Alpha@1; beta@2", Render(style, Cite("one", 1, "a"), Cite("two", 2, "b"), Cite("three", 3, "b", "a"))[2]);
    }

    [Theory]
    [InlineData(false, CslOutputFormat.PlainText, "Alpha")]
    [InlineData(false, CslOutputFormat.Html, "Alpha")]
    [InlineData(true, CslOutputFormat.PlainText, "see alpha")]
    [InlineData(true, CslOutputFormat.Html, "see alpha")]
    public void TheFirstVisibleItemOwnsNoteOpeningAndLiteralPrefixes(bool visiblePrefix, CslOutputFormat format, string expected) {
        BibliographyDocument data = BibliographyDocument.Parse("[{\"id\":\"a\",\"type\":\"book\"},{\"id\":\"b\",\"type\":\"book\",\"title\":\"alpha\"}]", BibliographyFormat.CslJson).Document;
        CslCitation cite = Cite("one", 1, "a", "b");
        cite.Items[visiblePrefix ? 1 : 0].Prefix = "see ";
        Assert.Equal(expected, new CslProcessor(data, Style("<text variable=\"title\"/>"), new CslRenderOptions { OutputFormat = format }).Render(new[] { cite }).Citations[0].Content);
    }

    [Fact]
    public void CollapseReevaluationRetainsTheCurrentNoteForBackreferences() {
        BibliographyDocument data = BibliographyDocument.Parse("[{\"id\":\"a\",\"type\":\"book\",\"author\":[{\"family\":\"Doe\"}],\"issued\":{\"date-parts\":[[2020]]}}]", BibliographyFormat.CslJson).Document;
        CslStyle style = Style("<names variable=\"author\"><name form=\"short\" suffix=\" \"/></names><date variable=\"issued\"><date-part name=\"year\"/></date><text variable=\"first-reference-note-number\" prefix=\"@\"/>", "collapse=\"year\" cite-group-delimiter=\"; \"");
        Assert.Equal("Doe 2020; 2020", new CslProcessor(data, style).Render(new[] { Cite("one", 1, "a", "a") }).Citations[0].Content);
    }

    [Fact]
    public void HostCanMarkTextBeforeANoteCitationAndTheProcessorCopiesTheFlag() {
        CslCitation cite = Cite("one", 1, "a"); cite.NoteHasPrecedingText = true;
        var processor = new CslProcessor(Data, Style("<text variable=\"title\"/>"));
        Assert.Equal("alpha", processor.Render(new[] { cite }).Citations[0].Content);
        cite.NoteHasPrecedingText = false;
        Assert.Equal("Alpha", processor.Render(new[] { cite }).Citations[0].Content);
    }
}
