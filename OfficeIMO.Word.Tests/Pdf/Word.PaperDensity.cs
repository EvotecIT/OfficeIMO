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
            TableFontSize = global::OfficeIMO.OfficeLength.Points(9), TableCellVerticalPadding = global::OfficeIMO.OfficeLength.Points(1.5),
            TableCellHorizontalPadding = global::OfficeIMO.OfficeLength.Points(3.5), TableParagraphSpacing = global::OfficeIMO.OfficeLength.Points(0), TableLineSpacing = global::OfficeIMO.OfficeLength.Points(11)
        } : new WordPaperStyle();
        var table = style.AddKeyValueTable(document, new[] { new KeyValuePair<string, string>("Host", "mail.example.test") });
        Assert.Equal(fontSize, table.Rows[1].Cells[1].Paragraphs[0].FontSize?.ToPoints());
        Assert.Equal(global::OfficeIMO.OfficeLength.Points((vertical) / 20D), table.StyleDetails!.MarginDefaultTop);
        Assert.Equal(global::OfficeIMO.OfficeLength.Points((vertical) / 20D), table.StyleDetails.MarginDefaultBottom);
        Assert.Equal(global::OfficeIMO.OfficeLength.Points((horizontal) / 20D), table.StyleDetails.MarginDefaultLeft);
        Assert.Equal(global::OfficeIMO.OfficeLength.Points((horizontal) / 20D), table.StyleDetails.MarginDefaultRight);
        Assert.Equal(spacing, table.Rows[1].Cells[1].Paragraphs[0].LineSpacingAfter?.ToPoints());
        Assert.Equal(spacing, table.Rows[1].Cells[1].Paragraphs[0].LineSpacingBefore?.ToPoints());
        if (compact) {
            Assert.Equal(11, table.Rows[1].Cells[1].Paragraphs[0].LineSpacing?.Height?.ToPoints());
            Assert.Equal(WordLineSpacingRule.Exact, table.Rows[1].Cells[1].Paragraphs[0].NativeLineSpacingRule);
        } else {
            Assert.Null(table.Rows[1].Cells[1].Paragraphs[0].LineSpacing?.Height?.ToPoints());
        }
    }
}
