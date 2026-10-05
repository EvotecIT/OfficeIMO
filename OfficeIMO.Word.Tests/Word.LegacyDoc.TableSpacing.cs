using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.LegacyDoc.Model;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LegacyDoc_TableSpacing_DoesNotInsertBlankBeforeFollowingBodyParagraph(bool authoredBlank) {
        using WordDocument document = WordDocument.Create();
        document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0].Text = "Cell";
        if (authoredBlank) document.AddParagraph();
        document.AddParagraph("After");
        byte[] bytes = document.ToBytes(WordFileFormat.Doc);

        // The main story is also consumed by Word; self-round-trip projection alone hid the extra paragraph.
        Assert.Equal("Cell\a\a" + (authoredBlank ? "\r" : "") + "After\r", ReadNativeTableSpacingStory(bytes));
        using WordDocument reopened = WordDocument.Load(new MemoryStream(bytes));
        Assert.Single(reopened.Tables);
        Assert.Equal(authoredBlank ? new[] { "", "After" } : new[] { "After" },
            reopened._wordprocessingDocument!.MainDocumentPart!.Document.Body!.Elements<Paragraph>().Select(p => p.InnerText));
        Assert.Equal("Cell", reopened.Tables[0].Rows[0].Cells[0].Paragraphs[0].Text);
    }

    [Fact]
    public void LegacyDoc_TableSpacing_KeepsAdjacentTablesDistinctAndFinalBodyTerminator() {
        using WordDocument document = WordDocument.Create();
        document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0].Text = "First";
        document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0].Text = "Second";
        byte[] bytes = document.ToBytes(WordFileFormat.Doc);
        Assert.Equal("First\a\a\rSecond\a\a\r", ReadNativeTableSpacingStory(bytes));
        using WordDocument reopened = WordDocument.Load(new MemoryStream(bytes));
        Assert.Equal(2, reopened.Tables.Count);
        Assert.Equal("First", reopened.Tables[0].Rows[0].Cells[0].Paragraphs[0].Text);
        Assert.Equal("Second", reopened.Tables[1].Rows[0].Cells[0].Paragraphs[0].Text);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void LegacyDoc_TableSpacing_PreservesGridParagraphDefaultsAndDirectOverride(bool derivedStyle, bool directOverride) {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(1, 1, WordTableStyle.TableGrid);
        if (derivedStyle) {
            var style = new Style { Type = StyleValues.Table, StyleId = "DerivedGrid", CustomStyle = true };
            style.Append(new StyleName { Val = "Derived Grid" });
            style.Append(new BasedOn { Val = "TableGrid" });
            document._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(style);
            table._tableProperties!.TableStyle = new TableStyle { Val = "DerivedGrid" };
        }
        WordParagraph paragraph = table.Rows[0].Cells[0].Paragraphs[0];
        paragraph.Text = "Grid";
        if (directOverride) paragraph.LineSpacingAfter = 120;

        using WordDocument reopened = WordDocument.Load(new MemoryStream(document.ToBytes(WordFileFormat.Doc)));
        WordParagraph restored = Assert.Single(Assert.Single(reopened.Tables).Rows[0].Cells[0].Paragraphs);
        Assert.Equal(directOverride ? 120 : 0, restored.LineSpacingAfter);
        Assert.Equal(240, restored.LineSpacing);
        Assert.Equal(WordLineSpacingRule.Auto, restored.LineSpacingRule);
    }

    private static string ReadNativeTableSpacingStory(byte[] bytes) {
        LegacyDocDocument model = LegacyDocDocument.Load(bytes);
        var streams = model.SourceCompoundFile!.Streams;
        byte[] word = streams["WordDocument"];
        Assert.True(LegacyDocFib.TryRead(word, out LegacyDocFib fib, out string? fibError), fibError);
        byte[] table = streams[fib.UsesOneTableStream ? "1Table" : "0Table"];
        Assert.True(LegacyDocPieceTable.TryRead(word, table, fib, 100000, out LegacyDocTextContent content, out string? textError), textError);
        return content.Text.Substring(0, fib.CcpText);
    }

    [Theory]
    [InlineData(false, false, false)]
    [InlineData(false, true, false)]
    [InlineData(false, true, true)]
    [InlineData(true, false, false)]
    [InlineData(true, true, false)]
    [InlineData(true, true, true)]
    public void LegacyDoc_TableSpacing_ParagraphStylePrecedesTableStyle(bool customTable, bool inheritedParagraph, bool directOverride) {
        using WordDocument document = WordDocument.Create();
        var styles = document._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        var paragraphStyle = new Style { Type = StyleValues.Paragraph, StyleId = "CellSpacing", CustomStyle = true };
        paragraphStyle.Append(new StyleName { Val = "Cell Spacing" }, new BasedOn { Val = "Normal" },
            new StyleParagraphProperties(new SpacingBetweenLines { After = "240", Line = "480", LineRule = LineSpacingRuleValues.Auto }));
        styles.Append(paragraphStyle);
        if (inheritedParagraph) {
            styles.Append(new Style(new StyleName { Val = "Derived Cell" }, new BasedOn { Val = "CellSpacing" }) {
                Type = StyleValues.Paragraph, StyleId = "DerivedCell", CustomStyle = true
            });
        }
        WordTable table = document.AddTable(1, 1, WordTableStyle.TableGrid);
        if (customTable) {
            styles.Append(new Style(new StyleName { Val = "Custom Cell Table" }, new BasedOn { Val = "TableNormal" },
                new StyleParagraphProperties(new SpacingBetweenLines { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto })) {
                Type = StyleValues.Table, StyleId = "CustomCellTable", CustomStyle = true
            });
            table._tableProperties!.TableStyle = new TableStyle { Val = "CustomCellTable" };
        }
        WordParagraph paragraph = table.Rows[0].Cells[0].Paragraphs[0];
        paragraph.Text = "Styled cell";
        paragraph._paragraph.ParagraphProperties ??= new ParagraphProperties();
        paragraph._paragraph.ParagraphProperties.ParagraphStyleId = new ParagraphStyleId { Val = inheritedParagraph ? "DerivedCell" : "CellSpacing" };
        if (directOverride) {
            paragraph.LineSpacingAfter = 120;
            paragraph.LineSpacing = 360;
            paragraph.LineSpacingRule = WordLineSpacingRule.Auto;
        }
        using WordDocument reopened = WordDocument.Load(new MemoryStream(document.ToBytes(WordFileFormat.Doc)));
        WordParagraph restored = Assert.Single(Assert.Single(reopened.Tables).Rows[0].Cells[0].Paragraphs);
        Assert.Equal(directOverride ? 120 : 240, restored.LineSpacingAfter);
        Assert.Equal(directOverride ? 360 : 480, restored.LineSpacing);
        Assert.Equal(WordLineSpacingRule.Auto, restored.LineSpacingRule);
        Assert.Equal(inheritedParagraph ? "LegacyDocDerivedCell" : "LegacyDocCellSpacing", restored.StyleId);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public void LegacyDoc_TableSpacing_AdjacentTableBookmarksExcludeSeparator(int contentControlDepth) {
        using WordDocument document = WordDocument.Create();
        WordTable first = document.AddTable(1, 1);
        first.Rows[0].Cells[0].Paragraphs[0].Text = "First";
        WordTable second = document.AddTable(1, 1);
        second.Rows[0].Cells[0].Paragraphs[0].Text = "Second";
        Body body = document._wordprocessingDocument!.MainDocumentPart!.Document.Body!;
        body.InsertBefore(new BookmarkStart { Id = "61", Name = "FirstTable" }, first._table);
        body.InsertAfter(new BookmarkEnd { Id = "61" }, first._table);
        body.InsertBefore(new BookmarkStart { Id = "62", Name = "SecondTable" }, second._table);
        body.InsertBefore(new BookmarkStart { Id = "63", Name = "BeforeSecond" }, second._table);
        body.InsertBefore(new BookmarkEnd { Id = "63" }, second._table);
        body.InsertAfter(new BookmarkEnd { Id = "62" }, second._table);
        if (contentControlDepth == 3) {
            var content = new SdtContentBlock();
            foreach (var marker in body.ChildElements.Where(child => child is BookmarkStart start && start.Id?.Value != "61" ||
                child is BookmarkEnd end && end.Id?.Value == "63").ToArray()) {
                marker.Remove();
                content.Append(marker);
            }
            body.InsertBefore(new SdtBlock(new SdtProperties(new SdtAlias { Val = "Leading markers" }), content), second._table);
        }
        for (int depth = 0; depth < (contentControlDepth == 3 ? 0 : contentControlDepth); depth++) {
            var content = new SdtContentBlock();
            while (body.FirstChild is not null && body.FirstChild is not SectionProperties) {
                var child = body.FirstChild;
                child.Remove();
                content.Append(child);
            }
            body.PrependChild(new SdtBlock(new SdtProperties(new SdtAlias { Val = "Tables" }), content));
        }
        byte[] bytes = document.ToBytes(WordFileFormat.Doc);
        string story = ReadNativeTableSpacingStory(bytes);
        LegacyDocDocument model = LegacyDocDocument.Load(bytes);
        int secondStart = story.IndexOf("Second", System.StringComparison.Ordinal);
        Assert.Equal(secondStart, Assert.Single(model.Bookmarks, b => b.Name == "SecondTable").StartCharacter);
        var point = Assert.Single(model.Bookmarks, b => b.Name == "BeforeSecond");
        Assert.Equal(secondStart, point.StartCharacter);
        Assert.Equal(secondStart, point.EndCharacter);
        Assert.Equal("First\a\a".Length, Assert.Single(model.Bookmarks, b => b.Name == "FirstTable").EndCharacter);
        using WordDocument reopened = WordDocument.Load(new MemoryStream(bytes));
        Assert.Equal(2, reopened.Tables.Count);
        Assert.Contains(reopened.Bookmarks, b => b.Name == "SecondTable");
        Assert.Contains(reopened.Tables[1].Rows[0].Cells[0].Paragraphs, p => p.Text == "Second");
    }
}
