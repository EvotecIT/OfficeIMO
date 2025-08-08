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
            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            var table = doc.Tables[0];
            Assert.True(table.RepeatHeaderRowAtTheTopOfEachPage);
            Assert.Equal("ff0000", table.Rows[0].Cells[0].ShadingFillColorHex);
            Assert.Equal(new List<int> { 1000, 4000 }, table.ColumnWidth);
            Assert.Equal(TableWidthUnitValues.Pct, table.ColumnWidthType);
            Assert.Equal("F1", table.Rows[table.Rows.Count - 1].Cells[0].Paragraphs[0]._paragraph.InnerText);
        }
    }
}
