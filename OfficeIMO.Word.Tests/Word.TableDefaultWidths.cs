using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(2)]
    [InlineData(6)]
    public void AddTable_DefaultColumnsFitTheTextArea(int columns) {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(2, columns);
        WordSection section = document.Sections[0];
        long textWidth = (section.PageSettings.Width!.Value - section.Margins.Left!.Value - section.Margins.Right!.Value).ToTwips();

        long total = table.GridColumnWidth.Sum(width => width.ToTwips());
        Assert.InRange(total, textWidth - columns, textWidth);
        Assert.All(table.Rows[0].Cells, cell => Assert.Equal(WordTableWidth.Length(table.GridColumnWidth[0]), cell.Width));
    }
}
