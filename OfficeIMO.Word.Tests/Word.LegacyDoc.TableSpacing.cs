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
}
