using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void HtmlToWord_TableBorderAttribute_AndCellContent() {
            string html = "<table border=\"2\"><tr><td>A1</td><td style=\"border:1px solid #ff0000\">B1</td></tr></table>";
            using WordDocument doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWord();
            var table = doc.Tables[0];

            var (style, size, color) = table.StyleDetails!.GetBorderProperties(WordTableBorderSide.Top);
            Assert.Equal(WordBorderStyle.Single, style);
            Assert.Equal(global::OfficeIMO.OfficeLength.Points(1.5), size);
            Assert.Equal(OfficeColor.Parse("000000"), color);

            Assert.Contains("A1", string.Join(string.Empty, table.Rows[0].Cells[0].Paragraphs.ConvertAll(p => p.Text)));
            Assert.Contains("B1", string.Join(string.Empty, table.Rows[0].Cells[1].Paragraphs.ConvertAll(p => p.Text)));

            var cell = table.Rows[0].Cells[1];
            Assert.Equal(WordBorderStyle.Single, cell.Borders.TopStyle);
            Assert.Equal(OfficeColor.Parse("FF0000"), cell.Borders.TopColor);
            Assert.Equal(global::OfficeIMO.OfficeLength.Points(0.75), cell.Borders.TopSize);
        }
    }
}

