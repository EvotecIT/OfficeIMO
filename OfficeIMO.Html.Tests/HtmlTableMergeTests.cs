using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests {
    public class HtmlTableMergeTests {
        [Fact]
        public void HtmlToWordConverter_HandlesRowspanAndColspan() {
            string html = "<table><tr><td rowspan=\"2\">A1</td><td>B1</td></tr><tr><td>B2</td></tr><tr><td colspan=\"2\">A3</td></tr></table>";
            using WordDocument doc = html.LoadFromHtml();
            var table = doc.Tables[0];

            Assert.Equal(MergedCellValues.Restart, table.Rows[0].Cells[0].VerticalMerge);
            Assert.Equal(MergedCellValues.Continue, table.Rows[1].Cells[0].VerticalMerge);
            Assert.Equal(MergedCellValues.Restart, table.Rows[2].Cells[0].HorizontalMerge);
            Assert.Equal(MergedCellValues.Continue, table.Rows[2].Cells[1].HorizontalMerge);
        }

        [Fact]
        public void WordToHtmlConverter_EmitsSpanAttributes() {
            using WordDocument doc = WordDocument.Create();
            WordTable table = doc.AddTable(3, 2);
            table.Rows[0].Cells[0].Paragraphs[0].Text = "A1";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "B1";
            table.Rows[1].Cells[1].Paragraphs[0].Text = "B2";
            table.Rows[2].Cells[0].Paragraphs[0].Text = "A3";

            table.MergeCells(0, 0, 2, 1);
            table.MergeCells(2, 0, 1, 2);

            string html = doc.ToHtml();
            Assert.Contains("rowspan=\"2\"", html);
            Assert.Contains("colspan=\"2\"", html);
        }

        [Fact]
        public void HtmlToWordConverter_RowSpanZero_StopsAtSectionEnd() {
            string html = "<table><tbody><tr><td rowspan=\"0\">A</td><td>B1</td></tr><tr><td>B2</td></tr></tbody><tfoot><tr><td>F1</td><td>F2</td></tr></tfoot></table>";
            using WordDocument doc = html.LoadFromHtml();
            var table = doc.Tables[0];

            Assert.Equal(MergedCellValues.Restart, table.Rows[0].Cells[0].VerticalMerge);
            Assert.Equal(MergedCellValues.Continue, table.Rows[1].Cells[0].VerticalMerge);
            Assert.Null(table.Rows[2].Cells[0].VerticalMerge);
        }
    }
}

