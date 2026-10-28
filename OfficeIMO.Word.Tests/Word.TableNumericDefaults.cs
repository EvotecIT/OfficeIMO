using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("1250")]
    [InlineData("1,250.50")]
    [InlineData("-25.5")]
    [InlineData("42%")]
    public void ModernTablePlainNumericValuesSaveWithThemeAlignment(string value) {
        using WordDocument source = WordDocument.Create();
        WordTable table = source.AddTable(2, 2);
        table.Rows[0].Cells[0].Paragraphs[0].Text = "Revenue";
        table.Rows[1].Cells[0].Paragraphs[0].Text = value;
        table.Rows[1].Cells[1].Paragraphs[0].Text = "Reported";
        using WordDocument loaded = WordDocument.Load(new MemoryStream(source.ToBytes()));
        Assert.Equal(WordParagraphAlignment.Right, loaded.Tables[0].Rows[1].Cells[0].Paragraphs[0].ParagraphAlignment);
        Assert.Null(loaded.Tables[0].Rows[1].Cells[1].Paragraphs[0].ParagraphAlignment);
        Assert.Equal(value, loaded.Tables[0].Rows[1].Cells[0].Paragraphs[0].Text);
    }

    [Fact]
    public void NumericDefaultsPreserveExplicitAlignmentStylesAndRichCellContent() {
        using WordDocument document = WordDocument.Create();
        var modern = document.AddTable(3, 2);
        var explicitAlignment = modern.Rows[1].Cells[0].Paragraphs[0];
        explicitAlignment.ParagraphAlignment = WordParagraphAlignment.Center;
        explicitAlignment.Text = "1250";
        var rich = modern.Rows[1].Cells[1].Paragraphs[0];
        rich.AddText("Label ");
        rich.AddText("1250");
        var styled = modern.Rows[2].Cells[0].Paragraphs[0];
        styled.Style = WordParagraphStyles.Heading1;
        styled.Text = "1250";
        var plain = document.AddTable(1, 1, WordTableStyle.TableGrid).Rows[0].Cells[0].Paragraphs[0];
        plain.Text = "1250";
        using WordDocument loaded = WordDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.Equal(WordParagraphAlignment.Center, loaded.Tables[0].Rows[1].Cells[0].Paragraphs[0].ParagraphAlignment);
        Assert.Null(loaded.Tables[0].Rows[1].Cells[1].Paragraphs[0].ParagraphAlignment);
        Assert.Null(loaded.Tables[0].Rows[2].Cells[0].Paragraphs[0].ParagraphAlignment);
        Assert.Null(loaded.Tables[1].Rows[0].Cells[0].Paragraphs[0].ParagraphAlignment);
        Assert.Empty(loaded.ValidateDocument());
    }
}
