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
        long textWidth = section.PageSettings.Width!.Value - (long)section.Margins.Left - section.Margins.Right;

        int total = table.GridColumnWidth.Sum();
        Assert.InRange(total, textWidth - columns, textWidth);
        Assert.All(table.Rows[0].Cells, cell => Assert.Equal(table.GridColumnWidth[0], cell.Width));
    }
}
