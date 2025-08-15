using OfficeIMO.Word.Pdf;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests.Pdf {
    public class TableLayoutCacheTests {
        [Fact]
        public void ColumnWidthsAreCachedPerTable() {
            using WordDocument document = WordDocument.Create();
            WordTable table = document.AddTable(2, 2);
            table.Rows[0].Cells[0].Width = 1440;
            table.Rows[1].Cells[0].Width = 1440;

            TableLayout first = TableLayoutCache.GetLayout(table);
            TableLayout second = TableLayoutCache.GetLayout(table);

            Assert.Same(first, second);
            Assert.Equal(2, first.ColumnWidths.Length);
            Assert.Equal(72f, first.ColumnWidths[0]);
        }

        [Fact]
        public void NestedTableWidthsPropagateToParent() {
            using WordDocument document = WordDocument.Create();
            WordTable outer = document.AddTable(1, 2);
            outer.Rows[0].Cells[0].Width = 1440;
            outer.Rows[0].Cells[1].Width = 720;

            WordTable inner = outer.Rows[0].Cells[1].AddTable(1, 1);
            inner.Rows[0].Cells[0].Width = 2880;

            TableLayout outerLayout = TableLayoutCache.GetLayout(outer);
            Assert.Equal(144f, outerLayout.ColumnWidths[1]);

            TableLayout innerLayout = TableLayoutCache.GetLayout(inner);
            TableLayout innerLayoutSecond = TableLayoutCache.GetLayout(inner);
            Assert.Same(innerLayout, innerLayoutSecond);
            Assert.Equal(144f, innerLayout.ColumnWidths[0]);
        }

        [Fact]
        public void RecursiveNestedTablesAreMeasured() {
            using WordDocument document = WordDocument.Create();
            WordTable outer = document.AddTable(1, 1);
            outer.Rows[0].Cells[0].Width = 720;
            WordTable middle = outer.Rows[0].Cells[0].AddTable(1, 1);
            middle.Rows[0].Cells[0].Width = 720;
            WordTable inner = middle.Rows[0].Cells[0].AddTable(1, 1);
            inner.Rows[0].Cells[0].Width = 2880;

            TableLayout outerLayout = TableLayoutCache.GetLayout(outer);
            Assert.Equal(144f, outerLayout.ColumnWidths[0]);

            TableLayout middleLayout = TableLayoutCache.GetLayout(middle);
            Assert.Equal(144f, middleLayout.ColumnWidths[0]);
            TableLayout innerLayout = TableLayoutCache.GetLayout(inner);
            Assert.Equal(144f, innerLayout.ColumnWidths[0]);
        }
    }
}

