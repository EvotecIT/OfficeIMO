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
            var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument(new HtmlToWordOptions());
            var table = doc.Tables[0];
            var (style, size, colorHex) = table.StyleDetails!.GetBorderProperties(WordTableBorderSide.Top);
            Assert.Equal(BorderValues.Single, style);
            Assert.Equal((UInt32Value)12U, size);
            Assert.Equal("FF0000", colorHex);
            var cell = table.Rows[0].Cells[0];
            Assert.Equal("00FF00", cell.ShadingFillColorHex);
            Assert.Equal(BorderValues.Dashed, cell.Borders.TopStyle);
            Assert.Equal("0000FF", cell.Borders.TopColorHex);
        }

        [Fact]
        public void HtmlToWord_TableStyles_MapsSideBorders() {
            string html = "<table style=\"border-left:3px double #112233\"><tr style=\"border-bottom:2px dotted #445566\"><td style=\"border-right:1px dashed #778899\">Cell</td></tr></table>";
            var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument(new HtmlToWordOptions());

            var table = doc.Tables[0];
            var (style, size, colorHex) = table.StyleDetails!.GetBorderProperties(WordTableBorderSide.Left);
            Assert.Equal(BorderValues.Double, style);
            Assert.Equal((UInt32Value)18U, size);
            Assert.Equal("112233", colorHex);
            var cell = table.Rows[0].Cells[0];
            Assert.Equal(BorderValues.Dotted, cell.Borders.BottomStyle);
            Assert.Equal((UInt32Value)12U, cell.Borders.BottomSize);
            Assert.Equal("445566", cell.Borders.BottomColorHex);
            Assert.Equal(BorderValues.Dashed, cell.Borders.RightStyle);
            Assert.Equal((UInt32Value)6U, cell.Borders.RightSize);
            Assert.Equal("778899", cell.Borders.RightColorHex);
        }

        [Fact]
        public void HtmlToWord_TableStyles_TextAlign() {
            string html = "<table><tr><td style=\"text-align:center\">One</td><td style=\"text-align:right\">Two</td></tr></table>";
            using var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument(new HtmlToWordOptions());
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
            using var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument(new HtmlToWordOptions());
            var table = doc.Tables[0];
            var cellLeft = table.Rows[0].Cells[0];
            var cellJustify = table.Rows[0].Cells[1];
            Assert.Equal(JustificationValues.Left, cellLeft.Paragraphs[0].ParagraphAlignment);
            Assert.Equal(JustificationValues.Both, cellJustify.Paragraphs[0].ParagraphAlignment);
            string back = doc.ToHtml();
            Assert.Contains("text-align:left", back, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("text-align:justify", back, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void HtmlToWord_TableStyles_WidthAuto() {
            string html = "<table style=\"width:auto\"><tr><td>Cell</td></tr></table>";
            using var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument(new HtmlToWordOptions());
            var table = doc.Tables[0];
            Assert.Equal(TableWidthUnitValues.Auto, table.WidthType);
            Assert.Equal(0, table.Width);
        }

        [Fact]
        public void HtmlToWord_TableStyles_MarginAutoCentersTable() {
            string html = "<table style=\"width:50%;margin:0 auto\"><tr><td>Cell</td></tr></table>";
            using var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument(new HtmlToWordOptions());
            var table = doc.Tables[0];

            Assert.Equal(TableRowAlignmentValues.Center, table.Alignment);
        }

        [Fact]
        public void HtmlToWord_TableStyles_AutoLeftMarginAlignsRight() {
            string html = "<table style=\"margin-left:auto;margin-right:0\"><tr><td>Cell</td></tr></table>";
            using var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument(new HtmlToWordOptions());
            var table = doc.Tables[0];

            Assert.Equal(TableRowAlignmentValues.Right, table.Alignment);
        }

        [Fact]
        public void HtmlToWord_TableStyles_AlignAttributeSetsTableAlignment() {
            string html = "<table align=\"center\"><tr><td>Cell</td></tr></table>";
            using var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument(new HtmlToWordOptions());
            var table = doc.Tables[0];

            Assert.Equal(TableRowAlignmentValues.Center, table.Alignment);
        }

        [Fact]
        public void HtmlToWord_TableStyles_CellSpacingAttributeSetsTableSpacing() {
            string html = "<table cellspacing=\"6\"><tr><td>One</td><td>Two</td></tr></table>";
            using var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument(new HtmlToWordOptions());
            var table = doc.Tables[0];

            Assert.Equal((short)90, table.StyleDetails!.CellSpacing);
        }

        [Fact]
        public void HtmlToWord_TableStyles_BorderSpacingSetsTableSpacing() {
            string html = "<table cellspacing=\"2\" style=\"border-spacing:5pt 10pt\"><tr><td>One</td><td>Two</td></tr></table>";
            using var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument(new HtmlToWordOptions());
            var table = doc.Tables[0];

            Assert.Equal((short)100, table.StyleDetails!.CellSpacing);
        }

        [Fact]
        public void HtmlToWord_TableStyles_VerticalAlignSetsCellVerticalAlignment() {
            string html = "<table><tr><td style=\"vertical-align:middle\">Middle</td><td valign=\"bottom\">Bottom</td></tr></table>";
            using var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument(new HtmlToWordOptions());
            var table = doc.Tables[0];

            Assert.Equal(TableVerticalAlignmentValues.Center, table.Rows[0].Cells[0].VerticalAlignment);
            Assert.Equal(TableVerticalAlignmentValues.Bottom, table.Rows[0].Cells[1].VerticalAlignment);
        }
    }
}
