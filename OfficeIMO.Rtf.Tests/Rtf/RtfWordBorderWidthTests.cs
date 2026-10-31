using OfficeIMO.Rtf;
using OfficeIMO.Word;
using OfficeIMO.Word.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

/// <summary>RTF states border widths in twips and Word in eighths of a point.</summary>
public class RtfWordBorderWidthTests {
    [Fact]
    public void WordBordersWriteTwipsAndComeBackAtTheSameWidth() {
        using WordDocument word = WordDocument.Create();
        word.Borders.TopStyle = WordBorderStyle.Single;
        word.Borders.TopSize = OfficeLength.Points(1.5);
        WordParagraph paragraph = word.AddParagraph("Boxed");
        paragraph.Borders.TopStyle = WordBorderStyle.Single;
        paragraph.Borders.TopSize = OfficeLength.Points(0.75);
        WordTable table = word.AddTable(1, 1);
        table.Rows[0].Cells[0].Borders.TopStyle = WordBorderStyle.Single;
        table.Rows[0].Cells[0].Borders.TopSize = OfficeLength.Points(0.5);

        RtfDocument rtf = word.ToRtfDocument();

        Assert.Equal(30, rtf.PageSetup.PageBorders.Top.Width);
        Assert.Equal(15, rtf.Paragraphs.First(candidate => candidate.ToPlainText() == "Boxed").TopBorder.Width);
        Assert.Equal(10, rtf.Blocks.OfType<RtfTable>().Single().Rows[0].Cells[0].TopBorder.Width);

        using WordDocument roundTrip = RtfDocument.Read(word.ToRtf(new RtfWriteOptions { IncludeGenerator = false })).Document.ToWordDocument();
        Assert.Equal(OfficeLength.Points(1.5), roundTrip.Borders.TopSize);
        Assert.Equal(OfficeLength.Points(0.75), roundTrip.Paragraphs.First(candidate => candidate.Text == "Boxed").Borders.TopSize);
        Assert.Equal(OfficeLength.Points(0.5), roundTrip.Tables[0].Rows[0].Cells[0].Borders.TopSize);
    }
}
