using OfficeIMO.Word;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, 0)]
    [InlineData(true, 0)]
    [InlineData(false, 1)]
    [InlineData(true, 1)]
    [InlineData(false, 2)]
    [InlineData(true, 2)]
    public void LegacyDoc_HeaderFooterTablePreservesCellsAndFollowingParagraph(bool footer, int variant) {
        HeaderFooterValues type = variant switch { 1 => HeaderFooterValues.First, 2 => HeaderFooterValues.Even, _ => HeaderFooterValues.Default };
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("BODY");
        source.AddHeadersAndFooters();
        source.DifferentFirstPage = true;
        source.DifferentOddAndEvenPages = true;
        WordHeaderFooter story = footer ? RequireSectionFooter(source, 0, type) : RequireSectionHeader(source, 0, type);
        WordTable table = story.AddTable(2, 2, WordTableStyle.TableGrid);
        for (int row = 0; row < 2; row++) for (int column = 0; column < 2; column++) {
            WordParagraph paragraph = table.Rows[row].Cells[column].Paragraphs[0];
            paragraph.Text = $"CELL{row + 1}{column + 1}";
            paragraph.FontFamily = "Arial";
            paragraph.FontSize = 12;
        }
        story.AddParagraph("AFTERTABLE");

        using WordDocument restored = WordDocument.Load(new MemoryStream(source.ToBytes(WordFileFormat.Doc)));
        WordHeaderFooter restoredStory = footer ? RequireSectionFooter(restored, 0, type) : RequireSectionHeader(restored, 0, type);
        AssertHeaderFooterTableCells(restoredStory);
        Assert.Contains(restoredStory.Paragraphs, paragraph => paragraph.Text == "AFTERTABLE");
        Assert.Equal("BODY", Assert.Single(restored.Paragraphs).Text);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LegacyDoc_HeaderFooterTablePreservesEmptyAndAdjacentTables(bool footer) {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("BODY");source.AddHeadersAndFooters();
        WordHeaderFooter story = footer ? source.Footer.Default : source.Header.Default;
        story.AddTable(1, 2, WordTableStyle.TableNormal);
        WordTable second = story.AddTable(1, 1, WordTableStyle.TableNormal);
        second.Rows[0].Cells[0].Paragraphs[0].Text = "SECOND";
        using WordDocument restored = WordDocument.Load(new MemoryStream(source.ToBytes(WordFileFormat.Doc)));
        WordHeaderFooter restoredStory = footer ? restored.Footer.Default : restored.Header.Default;
        Assert.Equal(2, restoredStory.Tables.Count);
        Assert.Equal(2, Assert.Single(restoredStory.Tables[0].Rows).Cells.Count);
        Assert.Equal("SECOND", Assert.Single(restoredStory.Tables[1].Rows[0].Cells[0].Paragraphs).Text);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void LegacyDoc_HeaderFooterTableResolvesItsOwnHyperlinkRelationship(bool footer, bool contentControl) {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph().AddHyperLink("BODY", new Uri("https://officeimo.net/body"));
        source.AddHeadersAndFooters();
        WordHeaderFooter story = footer ? source.Footer.Default : source.Header.Default;
        WordTable table = story.AddTable(1, 1, WordTableStyle.TableNormal);
        Uri target = new Uri("https://officeimo.net/story-table");
        WordParagraph paragraph = table.Rows[0].Cells[0].Paragraphs[0];
        paragraph.AddHyperLink("TABLELINK", target);
        if (contentControl) {
            Hyperlink link = Assert.Single(paragraph._paragraph.Elements<Hyperlink>());
            link.Remove();
            paragraph._paragraph.Append(new SdtRun(new SdtContentRun(link)));
        }
        using WordDocument restored = WordDocument.Load(new MemoryStream(source.ToBytes(WordFileFormat.Doc)));
        WordHeaderFooter restoredStory = footer ? restored.Footer.Default : restored.Header.Default;
        Assert.Contains(Assert.Single(restoredStory.Tables).Rows[0].Cells[0].Paragraphs.SelectMany(paragraph => paragraph.GetRuns()),
            run => run.IsHyperLink && run.Hyperlink!.Uri == target);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void LegacyDoc_HeaderFooterTableResolvesItsOwnPictureRelationship(bool footer, bool nested) {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("BODY");source.AddHeadersAndFooters();
        WordHeaderFooter story = footer ? source.Footer.Default : source.Header.Default;
        WordTable table = story.AddTable(1, 1, WordTableStyle.TableNormal);
        WordTable target = nested ? table.Rows[0].Cells[0].AddTable(1, 1, WordTableStyle.TableNormal) : table;
        target.Rows[0].Cells[0].Paragraphs[0].AddImage(Path.Combine(AppContext.BaseDirectory, "Images", "EvotecLogo.png"), 18, 12);
        using WordDocument restored = WordDocument.Load(new MemoryStream(source.ToBytes(WordFileFormat.Doc)));
        WordHeaderFooter restoredStory = footer ? restored.Footer.Default : restored.Header.Default;
        WordTable restoredTable = Assert.Single(restoredStory.Tables);
        WordTable restoredTarget = nested ? Assert.Single(restoredTable.Rows[0].Cells[0].NestedTables) : restoredTable;
        WordParagraph run = Assert.Single(restoredTarget.Rows[0].Cells[0].Paragraphs.SelectMany(paragraph => paragraph.GetRuns()), paragraph => paragraph.IsImage);
        Assert.Equal(18, run.Image!.Width!.Value, 3);
        Assert.Equal(12, run.Image.Height!.Value, 3);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LegacyDoc_HeaderFooterTablePreservesNestedCellContent(bool footer) {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("BODY");source.AddHeadersAndFooters();
        WordHeaderFooter story = footer ? source.Footer.Default : source.Header.Default;
        WordTable table = story.AddTable(1, 1, WordTableStyle.TableNormal);
        table.Rows[0].Cells[0].Paragraphs[0].Text = "OUTER";
        WordTable nested = table.Rows[0].Cells[0].AddTable(1, 1, WordTableStyle.TableNormal);
        nested.Rows[0].Cells[0].Paragraphs[0].Text = "NESTED";
        using WordDocument restored = WordDocument.Load(new MemoryStream(source.ToBytes(WordFileFormat.Doc)));
        WordHeaderFooter restoredStory = footer ? restored.Footer.Default : restored.Header.Default;
        WordTable outer = Assert.Single(restoredStory.Tables);
        Assert.Contains(outer.Rows[0].Cells[0].Paragraphs, paragraph => paragraph.Text == "OUTER");
        WordTable restoredNested = Assert.Single(outer.Rows[0].Cells[0].NestedTables);
        Assert.Equal("NESTED", Assert.Single(restoredNested.Rows[0].Cells[0].Paragraphs).Text);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LegacyDoc_HeaderFooterTablePreservesCellShadingAndBorderOverride(bool footer) {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("BODY");source.AddHeadersAndFooters();
        WordHeaderFooter story = footer ? source.Footer.Default : source.Header.Default;
        WordTableCell cell = story.AddTable(1, 1, WordTableStyle.TableGrid).Rows[0].Cells[0];
        cell.Paragraphs[0].Text = "FORMATTED";
        cell.ShadingFillColorHex = "FFFF00";
        cell.Borders.BottomStyle = WordBorderStyle.Double;
        cell.Borders.BottomColorHex = "00FF00";
        cell.Borders.BottomSize = 8;
        using WordDocument restored = WordDocument.Load(new MemoryStream(source.ToBytes(WordFileFormat.Doc)));
        WordHeaderFooter restoredStory = footer ? restored.Footer.Default : restored.Header.Default;
        WordTableCell restoredCell = Assert.Single(restoredStory.Tables).Rows[0].Cells[0];
        Assert.Equal("FFFF00", restoredCell.ShadingFillColorHex);
        Assert.Equal(WordBorderStyle.Double, restoredCell.Borders.BottomStyle);
        Assert.Equal("00FF00", restoredCell.Borders.BottomColorHex);
        Assert.Equal(8U, restoredCell.Borders.BottomSize);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LegacyDoc_HeaderFooterTablePreservesCellBookmarkRange(bool footer) {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("BODY");source.AddHeadersAndFooters();
        WordHeaderFooter story = footer ? source.Footer.Default : source.Header.Default;
        WordParagraph paragraph = story.AddTable(1, 1, WordTableStyle.TableNormal).Rows[0].Cells[0].Paragraphs[0];
        paragraph._paragraph.RemoveAllChildren<Run>();
        paragraph._paragraph.Append(new BookmarkStart { Id = "51", Name = "StoryCellBookmark" },
            new Run(new Text("MARKED")), new BookmarkEnd { Id = "51" });
        using WordDocument restored = WordDocument.Load(new MemoryStream(source.ToBytes(WordFileFormat.Doc)));
        WordHeaderFooter restoredStory = footer ? restored.Footer.Default : restored.Header.Default;
        Paragraph restoredParagraph = Assert.Single(Assert.Single(restoredStory.Tables).Rows[0].Cells[0].Paragraphs,
            candidate => candidate.Bookmark?.Name == "StoryCellBookmark")._paragraph;
        BookmarkStart start = Assert.Single(restoredParagraph.Elements<BookmarkStart>());
        Assert.Equal("StoryCellBookmark", start.Name!.Value);
        Assert.Equal(start.Id!.Value, Assert.Single(restoredParagraph.Elements<BookmarkEnd>()).Id!.Value);
        Assert.Equal("MARKED", restoredParagraph.InnerText);
    }

    [Theory]
    [InlineData(false, "WordHeaderTableNative.doc")]
    [InlineData(true, "WordFooterTableNative.doc")]
    public void LegacyDoc_ImportsIndependentWordHeaderFooterTable(bool footer, string file) {
        using WordDocument document = WordDocument.Load(Path.Combine(AppContext.BaseDirectory, "Documents", file));
        WordHeaderFooter story = footer ? document.Footer.Default : document.Header.Default;
        AssertHeaderFooterTableCells(story);
        Assert.Equal("AFTERTABLE", Assert.Single(story.Paragraphs).Text);
    }

    private static void AssertHeaderFooterTableCells(WordHeaderFooter story) {
        WordTable table = Assert.Single(story.Tables);
        Assert.Equal(2, table.Rows.Count);
        for (int row = 0; row < 2; row++) {
            Assert.Equal(2, table.Rows[row].Cells.Count);
            for (int column = 0; column < 2; column++) {
                Assert.Equal($"CELL{row + 1}{column + 1}", Assert.Single(table.Rows[row].Cells[column].Paragraphs).Text);
            }
        }
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void LegacyDoc_HeaderFooterTableRepeatedSavePreservesIntentionalParagraphs(bool footer, bool trailingBlank) {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("BODY");
        source.AddHeadersAndFooters();
        WordHeaderFooter story = footer ? source.Footer.Default : source.Header.Default;
        story.AddTable(1, 1, WordTableStyle.TableNormal).Rows[0].Cells[0].Paragraphs[0].Text = "CELL";
        story.AddParagraph("AFTER");
        if (trailingBlank) story.AddParagraph();
        string[] expected = trailingBlank ? new[] { "AFTER", "" } : new[] { "AFTER" };
        byte[] bytes = source.ToBytes(WordFileFormat.Doc);
        for (int cycle = 0; cycle < 3; cycle++) {
            using WordDocument restored = WordDocument.Load(new MemoryStream(bytes));
            WordHeaderFooter restoredStory = footer ? restored.Footer.Default : restored.Header.Default;
            Assert.Equal(expected, restoredStory.Paragraphs.Select(paragraph => paragraph.Text));
            Assert.Equal("CELL", Assert.Single(restoredStory.Tables).Rows[0].Cells[0].Paragraphs[0].Text);
            bytes = restored.ToBytes(WordFileFormat.Doc);
        }
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void LegacyDoc_TableEmptyCellBookmarkRemainsEditableAfterRepeatedSave(int storyKind) {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("BODY");
        source.AddHeadersAndFooters();
        WordTable table = storyKind == 0 ? source.AddTable(1, 2, WordTableStyle.TableNormal)
            : (storyKind == 1 ? (WordHeaderFooter)source.Header.Default : source.Footer.Default).AddTable(1, 2, WordTableStyle.TableNormal);
        table.Rows[0].Cells[0].Paragraphs[0].Text = "FIRST";
        table.Rows[0].Cells[1].Paragraphs[0].AddBookmark("EmptyCell");
        byte[] bytes = source.ToBytes(WordFileFormat.Doc);
        for (int cycle = 0; cycle < 3; cycle++) {
            using WordDocument restored = WordDocument.Load(new MemoryStream(bytes));
            WordTable restoredTable = storyKind == 0 ? Assert.Single(restored.Tables)
                : Assert.Single((storyKind == 1 ? (WordHeaderFooter)restored.Header.Default : restored.Footer.Default).Tables);
            WordParagraph empty = Assert.Single(restoredTable.Rows[0].Cells[1].Paragraphs);
            Assert.Equal("", empty.Text);
            Assert.Equal("EmptyCell", empty.Bookmark?.Name);
            Assert.Equal("FIRST", restoredTable.Rows[0].Cells[0].Paragraphs[0].Text);
            bytes = restored.ToBytes(WordFileFormat.Doc);
        }
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void LegacyDoc_NestedTableKeepsMultipleParagraphsWithinTheirCells(int storyKind) {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("BODY");
        source.AddHeadersAndFooters();
        WordTable outer = storyKind == 0 ? source.AddTable(1, 1, WordTableStyle.TableNormal)
            : (storyKind == 1 ? (WordHeaderFooter)source.Header.Default : source.Footer.Default).AddTable(1, 1, WordTableStyle.TableNormal);
        WordTable nested = outer.Rows[0].Cells[0].AddTable(2, 2, WordTableStyle.TableNormal);
        for (int row = 0; row < 2; row++) for (int column = 0; column < 2; column++) {
            WordTableCell cell = nested.Rows[row].Cells[column];
            cell.Paragraphs[0].Text = $"CELL{row}{column}";
            cell.AddParagraph("SECOND");
        }
        byte[] bytes = source.ToBytes(WordFileFormat.Doc);
        for (int cycle = 0; cycle < 2; cycle++) {
            using WordDocument restored = WordDocument.Load(new MemoryStream(bytes));
            WordTable restoredOuter = storyKind == 0 ? Assert.Single(restored.Tables)
                : Assert.Single((storyKind == 1 ? (WordHeaderFooter)restored.Header.Default : restored.Footer.Default).Tables);
            WordTable restoredNested = Assert.Single(restoredOuter.Rows[0].Cells[0].NestedTables);
            Assert.Equal(2, restoredNested.Rows.Count);
            for (int row = 0; row < 2; row++) {
                Assert.Equal(2, restoredNested.Rows[row].Cells.Count);
                for (int column = 0; column < 2; column++) Assert.Equal(new[] { $"CELL{row}{column}", "SECOND" },
                    restoredNested.Rows[row].Cells[column].Paragraphs.Select(paragraph => paragraph.Text));
            }
            bytes = restored.ToBytes(WordFileFormat.Doc);
        }
    }
}
