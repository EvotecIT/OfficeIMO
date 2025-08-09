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
            using WordDocument doc = html.LoadFromHtml();
            var table = doc.Tables[0];

            var (style, size, colorHex) = table.StyleDetails.GetBorderProperties(WordTableBorderSide.Top);
            Assert.Equal(BorderValues.Single, style);
            Assert.Equal((UInt32Value)12U, size);
            Assert.Equal("000000", colorHex);

            Assert.Contains("A1", string.Join(string.Empty, table.Rows[0].Cells[0].Paragraphs.ConvertAll(p => p.Text)));
            Assert.Contains("B1", string.Join(string.Empty, table.Rows[0].Cells[1].Paragraphs.ConvertAll(p => p.Text)));

            var cell = table.Rows[0].Cells[1];
            Assert.Equal(BorderValues.Single, cell.Borders.TopStyle);
            Assert.Equal("ff0000", cell.Borders.TopColorHex);
            Assert.Equal((UInt32Value)6U, cell.Borders.TopSize);
        }
    }
}

