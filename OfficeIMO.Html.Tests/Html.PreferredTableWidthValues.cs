using System.Linq;
using OfficeIMO.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Html {
    [Fact]
    public void HtmlToWord_MixedColumnPreferencesKeepUnitsAndMissingPositions() {
        using WordDocument document = HtmlConversionDocument.Parse("""
            <table style="width:400pt"><colgroup><col style="width:40pt"><col><col style="width:37.02%"><col style="width:auto"></colgroup>
            <tr><td>A</td><td>B</td><td>C</td><td>D</td></tr></table>
            """).ToWord();
        WordTable table = Assert.Single(document.Tables);
        Assert.Equal(WordTableWidth.Length(OfficeLength.Points(400)), table.Width);
        Assert.Equal(new WordTableWidth?[] {
            WordTableWidth.Length(OfficeLength.Points(40)), null, WordTableWidth.Percent(37.02), WordTableWidth.Auto
        }, table.ColumnWidth);
        Assert.Null(table.Rows[0].Cells[1]._tableCellProperties!.TableCellWidth);
        Assert.All(table.GridColumnWidth, width => Assert.True(width > OfficeLength.Zero));
    }

    [Fact]
    public void WordToHtml_ExportsCellAndTableWidthsByTheirOwnInterpretations() {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(1, 2);
        table.Width = WordTableWidth.Percent(37.02);
        table.Rows[0].Cells[0].Width = WordTableWidth.Length(OfficeLength.Points(72));
        table.Rows[0].Cells[1].Width = WordTableWidth.Percent(60);
        string html = document.ToHtml();
        Assert.Contains("37.02%", html);
        Assert.Contains("96px", html);
        Assert.Contains("60%", html);
    }
}
