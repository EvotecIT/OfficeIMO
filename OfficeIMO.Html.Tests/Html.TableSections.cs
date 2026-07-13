using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using System.Collections.Generic;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void HtmlToWord_TableSections_ColGroupAndHeader() {
            string html = "<table><colgroup><col style=\"width:20%\"><col style=\"width:80%\"></colgroup><thead><tr style=\"background-color:#ff0000\"><th>H1</th><th>H2</th></tr></thead><tbody><tr><td>B1</td><td>B2</td></tr></tbody><tfoot><tr><td>F1</td><td>F2</td></tr></tfoot></table>";
            var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument(new HtmlToWordOptions());
            var table = doc.Tables[0];
            Assert.True(table.RepeatHeaderRowAtTheTopOfEachPage);
            Assert.Equal("FF0000", table.Rows[0].Cells[0].ShadingFillColorHex);
            Assert.Equal(new List<int> { 1000, 4000 }, table.ColumnWidth);
            Assert.Equal(TableWidthUnitValues.Pct, table.ColumnWidthType);
            Assert.Equal("F1", table.Rows[table.Rows.Count - 1].Cells[0].Paragraphs[0]._paragraph.InnerText);
        }

        [Fact]
        public void HtmlToWord_TableSections_TheadMarksEveryHeaderRowRepeated() {
            string html = "<table><thead><tr><th>H1</th></tr><tr><th>H2</th></tr></thead><tbody><tr><td>B1</td></tr></tbody></table>";

            var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument(new HtmlToWordOptions());

            var table = doc.Tables[0];
            Assert.True(table.Rows[0].RepeatHeaderRowAtTheTopOfEachPage);
            Assert.True(table.Rows[1].RepeatHeaderRowAtTheTopOfEachPage);
            Assert.False(table.Rows[2].RepeatHeaderRowAtTheTopOfEachPage);
        }

        [Fact]
        public void HtmlToWord_TableSections_ColGroupDecimalPercent() {
            string html = "<table><colgroup><col style=\"width:13.5%\"><col style=\"width:86.5%\"></colgroup><tbody><tr><td>A</td><td>B</td></tr></tbody></table>";
            var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument(new HtmlToWordOptions());
            var table = doc.Tables[0];
            Assert.Equal(new List<int> { 675, 4325 }, table.ColumnWidth);
            Assert.Equal(TableWidthUnitValues.Pct, table.ColumnWidthType);
        }

        [Fact]
        public void HtmlToWord_TableSections_RowSpanReservesColumnsForFollowingRows() {
            string html = "<table><tbody><tr><td rowspan=\"2\">A</td><td>B</td></tr><tr><td>C</td><td>D</td></tr></tbody></table>";
            var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument(new HtmlToWordOptions());
            var table = doc.Tables[0];

            Assert.Equal(2, table.Rows.Count);
            Assert.Equal(3, table.Rows[0].Cells.Count);
            Assert.Equal("A", table.Rows[0].Cells[0].Paragraphs[0]._paragraph.InnerText);
            Assert.Equal("B", table.Rows[0].Cells[1].Paragraphs[0]._paragraph.InnerText);
            Assert.Equal("C", table.Rows[1].Cells[1].Paragraphs[0]._paragraph.InnerText);
            Assert.Equal("D", table.Rows[1].Cells[2].Paragraphs[0]._paragraph.InnerText);
        }
    }
}
