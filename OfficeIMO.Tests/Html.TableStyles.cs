using System;
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

        [Fact]
        public void HtmlToWord_TableStyles_TextAlign() {
            string html = "<table><tr><td style=\"text-align:center\">One</td><td style=\"text-align:right\">Two</td></tr></table>";
            using var doc = html.LoadFromHtml(new HtmlToWordOptions());
            var table = doc.Tables[0];
            var cell1 = table.Rows[0].Cells[0];
            var cell2 = table.Rows[0].Cells[1];
            Assert.Equal(JustificationValues.Center, cell1.Paragraphs[0].ParagraphAlignment);
            Assert.Equal(JustificationValues.Right, cell2.Paragraphs[0].ParagraphAlignment);
            string back = doc.ToHtml();
            Assert.Contains("text-align:center", back, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("text-align:right", back, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void HtmlToWord_TableStyles_TextAlign_LeftAndJustify() {
            string html = "<table><tr><td style=\"text-align:left\">L</td><td style=\"text-align:justify\">J</td></tr></table>";
            using var doc = html.LoadFromHtml(new HtmlToWordOptions());
            var table = doc.Tables[0];
            var cellLeft = table.Rows[0].Cells[0];
            var cellJustify = table.Rows[0].Cells[1];
            Assert.Equal(JustificationValues.Left, cellLeft.Paragraphs[0].ParagraphAlignment);
            Assert.Equal(JustificationValues.Both, cellJustify.Paragraphs[0].ParagraphAlignment);
            string back = doc.ToHtml();
            Assert.Contains("text-align:left", back, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("text-align:justify", back, StringComparison.OrdinalIgnoreCase);
        }
    }
}
