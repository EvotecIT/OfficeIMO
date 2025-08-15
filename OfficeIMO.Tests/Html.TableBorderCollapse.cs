using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void HtmlToWord_TableBorderCollapse_Collapse() {
            string html = "<table style=\"border-collapse:collapse;border:2px solid #ff0000\"><tr><td>A1</td><td>B1</td></tr></table>";
            using var doc = html.LoadFromHtml(new HtmlToWordOptions());
            var table = doc.Tables[0];
            var insideH = table.StyleDetails.GetBorderProperties(WordTableBorderSide.InsideHorizontal);
            Assert.Equal(BorderValues.Single, insideH.Style);
            Assert.Equal((UInt32Value)12U, insideH.Size);
            Assert.Equal("ff0000", insideH.ColorHex);
            var cell = table.Rows[0].Cells[0];
            Assert.Null(cell.Borders.TopStyle);
        }

        [Fact]
        public void HtmlToWord_TableBorderCollapse_Separate() {
            string html = "<table style=\"border-collapse:separate;border:2px solid #ff0000\"><tr><td>A1</td><td>B1</td></tr></table>";
            using var doc = html.LoadFromHtml(new HtmlToWordOptions());
            var table = doc.Tables[0];
            var insideH = table.StyleDetails.GetBorderProperties(WordTableBorderSide.InsideHorizontal);
            Assert.Null(insideH.Style);
            var cell = table.Rows[0].Cells[0];
            Assert.Equal(BorderValues.Single, cell.Borders.TopStyle);
            Assert.Equal((UInt32Value)12U, cell.Borders.TopSize);
            Assert.Equal("ff0000", cell.Borders.TopColorHex);
        }
    }
}

