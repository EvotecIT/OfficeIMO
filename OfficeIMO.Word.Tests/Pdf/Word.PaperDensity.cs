using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordPaperDensityTests {
    [Theory]
    [InlineData(false, 10, 90, 100, 2)]
    [InlineData(true, 9, 30, 70, 0)]
    public void TableDensityPreservesDefaultsAndAllowsCompactPaperTables(bool compact, double fontSize, int vertical, int horizontal, double spacing) {
        using WordDocument document = WordDocument.Create();
        var style = compact ? new WordPaperStyle {
            TableFontSizePoints = 9, TableCellVerticalPadding = 30,
            TableCellHorizontalPadding = 70, TableParagraphSpacingPoints = 0, TableLineSpacingPoints = 11
        } : new WordPaperStyle();
        var table = style.AddKeyValueTable(document, new[] { new KeyValuePair<string, string>("Host", "mail.example.test") });
        Assert.Equal(fontSize, table.Rows[1].Cells[1].Paragraphs[0].FontSizePoints);
        Assert.Equal((short)vertical, table.StyleDetails!.MarginDefaultTopWidth);
        Assert.Equal((short)vertical, table.StyleDetails.MarginDefaultBottomWidth);
        Assert.Equal((short)horizontal, table.StyleDetails.MarginDefaultLeftWidth);
        Assert.Equal((short)horizontal, table.StyleDetails.MarginDefaultRightWidth);
        Assert.Equal(spacing, table.Rows[1].Cells[1].Paragraphs[0].LineSpacingAfterPoints);
        Assert.Equal(spacing, table.Rows[1].Cells[1].Paragraphs[0].LineSpacingBeforePoints);
        if (compact) {
            Assert.Equal(11, table.Rows[1].Cells[1].Paragraphs[0].LineSpacingPoints);
            Assert.Equal(WordLineSpacingRule.Exact, table.Rows[1].Cells[1].Paragraphs[0].LineSpacingRule);
        } else {
            Assert.Null(table.Rows[1].Cells[1].Paragraphs[0].LineSpacingPoints);
        }
    }
}
