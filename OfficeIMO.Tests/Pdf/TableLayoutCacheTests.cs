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
    }
}

