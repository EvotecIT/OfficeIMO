using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void LegacyDoc_ListTabsImportedFromWordRetainAlignmentAndCustomPosition() {
        string path = Path.Combine(AppContext.BaseDirectory, "Documents", "Word", "LegacyLists", "WordListNumberTabs.doc");
        using WordDocument source = WordDocument.Load(path);
        AssertListNumberTab(source, 3000);
        using WordDocument docx = WordDocument.Load(new MemoryStream(source.ToBytes(WordFileFormat.Docx)));
        AssertListNumberTab(docx, 3000);
        using WordDocument native = WordDocument.Load(new MemoryStream(source.ToBytes(WordFileFormat.Doc)));
        AssertListNumberTab(native, 3000);
    }

    [Fact]
    public void LegacyDoc_ListTabsWrittenFromNumberingDefinitionsRemainEditable() {
        using WordDocument source = CreateNativeListDefinitionControl();
        Level level = source.Lists.Single().Numbering.Levels[0].OpenXmlElement;
        level.PreviousParagraphProperties ??= new PreviousParagraphProperties();
        level.PreviousParagraphProperties.Tabs = new Tabs(new TabStop {
            Val = TabStopValues.Number, Position = 3000, Leader = TabStopLeaderCharValues.None
        });
        using WordDocument restored = WordDocument.Load(new MemoryStream(source.ToBytes(WordFileFormat.Doc)));
        AssertListNumberTab(restored, 3000);
        TabStop tab = restored.Lists.Single().Numbering.Levels[0].OpenXmlElement
            .PreviousParagraphProperties!.Tabs!.Elements<TabStop>().Single();
        tab.Position = 2200;
        using WordDocument edited = WordDocument.Load(new MemoryStream(restored.ToBytes(WordFileFormat.Doc)));
        AssertListNumberTab(edited, 2200);
    }

    [Fact]
    public void LegacyDoc_ListTabsWrittenOutOfOrderKeepTheirPositionsAndLeaders() {
        using WordDocument source = WordDocument.Create();
        WordParagraph paragraph = source.AddParagraph("TABS");
        paragraph.AddTabStop(3000, WordTabAlignment.Number, WordTabLeader.Dot);
        paragraph.AddTabStop(720, WordTabAlignment.Left, WordTabLeader.Hyphen);
        using WordDocument restored = WordDocument.Load(new MemoryStream(source.ToBytes(WordFileFormat.Doc)));
        Assert.Collection(restored.Paragraphs.Single(p => p.Text == "TABS").TabStops,
            tab => {
                Assert.Equal(720, tab.Position);
                Assert.Equal(WordTabAlignment.Left, tab.Alignment);
                Assert.Equal(WordTabLeader.Hyphen, tab.Leader);
            },
            tab => {
                Assert.Equal(3000, tab.Position);
                Assert.Equal(WordTabAlignment.Number, tab.Alignment);
                Assert.Equal(WordTabLeader.Dot, tab.Leader);
            });
    }

    [Fact]
    public void LegacyDoc_ListTabsAtTheNativeCountLimitRemainWritable() {
        using WordDocument source = WordDocument.Create();
        WordParagraph paragraph = source.AddParagraph("TABS");
        for (int index = 64; index > 0; index--) paragraph.AddTabStop(index * 200, WordTabAlignment.Left);
        using WordDocument restored = WordDocument.Load(new MemoryStream(source.ToBytes(WordFileFormat.Doc)));
        Assert.Equal(64, restored.Paragraphs.Single(p => p.Text == "TABS").TabStops.Count);
        paragraph.AddTabStop(13000, WordTabAlignment.Left);
        NotSupportedException error = Assert.Throws<NotSupportedException>(() => source.ToBytes(WordFileFormat.Doc));
        Assert.Contains("64 added or cleared tab stops", error.Message);
        using WordDocument docx = WordDocument.Load(new MemoryStream(source.ToBytes(WordFileFormat.Docx)));
        Assert.Equal(65, docx.Paragraphs.Single(p => p.Text == "TABS").TabStops.Count);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public void LegacyDoc_ListTabsInParagraphsRetainPublicSettingsAcrossStories(int story) {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("BODY");
        if (story >= 2) source.AddHeadersAndFooters();
        WordParagraph paragraph = story switch {
            1 => source.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0],
            2 => source.Header.Default.AddParagraph("MARKED"),
            3 => source.Footer.Default.AddParagraph("MARKED"),
            _ => source.AddParagraph("MARKED")
        };
        paragraph.Text = "MARKED";
        paragraph.AddTabStop(2200, WordTabAlignment.Number, WordTabLeader.Dot);
        using WordDocument restored = WordDocument.Load(new MemoryStream(source.ToBytes(WordFileFormat.Doc)));
        WordParagraph actual = story switch {
            1 => restored.Tables.Single().Rows[0].Cells[0].Paragraphs[0],
            2 => restored.Header.Default.Paragraphs.Single(p => p.Text == "MARKED"),
            3 => restored.Footer.Default.Paragraphs.Single(p => p.Text == "MARKED"),
            _ => restored.Paragraphs.Single(p => p.Text == "MARKED")
        };
        WordTabStop tab = Assert.Single(actual.TabStops);
        Assert.Equal(WordTabAlignment.Number, tab.Alignment);
        Assert.Equal(WordTabLeader.Dot, tab.Leader);
        Assert.Equal(2200, tab.Position);
        tab.Position = 2600;
        using WordDocument edited = WordDocument.Load(new MemoryStream(restored.ToBytes(WordFileFormat.Docx)));
        WordParagraph projected = story switch {
            1 => edited.Tables.Single().Rows[0].Cells[0].Paragraphs[0],
            2 => edited.Header.Default.Paragraphs.Single(p => p.Text == "MARKED"),
            3 => edited.Footer.Default.Paragraphs.Single(p => p.Text == "MARKED"),
            _ => edited.Paragraphs.Single(p => p.Text == "MARKED")
        };
        Assert.Equal(2600, Assert.Single(projected.TabStops).Position);
        Assert.Equal(WordTabAlignment.Number, Assert.Single(projected.TabStops).Alignment);
    }

    private static void AssertListNumberTab(WordDocument document, int position) {
        PreviousParagraphProperties properties = Assert.IsType<PreviousParagraphProperties>(document.Lists.Single().Numbering.Levels[0].OpenXmlElement.PreviousParagraphProperties);
        Tabs tabs = Assert.IsType<Tabs>(properties.Tabs);
        TabStop tab = Assert.Single(tabs.Elements<TabStop>());
        Assert.Equal(TabStopValues.Number, tab.Val!.Value);
        Assert.Equal(position, tab.Position!.Value);
    }
}
