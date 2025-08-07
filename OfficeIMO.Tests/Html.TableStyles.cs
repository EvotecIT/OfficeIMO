using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void HtmlToWord_TableStyles_BorderAndShading() {
            string html = "<table style=\"border:2px solid #ff0000\"><tr style=\"background-color:#00ff00\"><td style=\"border:1px dashed #0000ff\">Cell</td></tr></table>";
            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            var table = doc.Tables[0];
            var (style, size, colorHex) = table.StyleDetails.GetBorderProperties(WordTableBorderSide.Top);
            Assert.Equal(BorderValues.Single, style);
            Assert.Equal((UInt32Value)12U, size);
            Assert.Equal("ff0000", colorHex);
            var cell = table.Rows[0].Cells[0];
            Assert.Equal("00ff00", cell.ShadingFillColorHex);
            Assert.Equal(BorderValues.Dashed, cell.Borders.TopStyle);
            Assert.Equal("0000ff", cell.Borders.TopColorHex);
        }
    }
}
