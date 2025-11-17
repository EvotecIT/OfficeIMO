using System.Linq;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public class WordTableParentEqualityTests {
        [Fact]
        public void TableRowAndCellParentEquality_WorkAsExpected() {
            using var doc = WordDocument.Create();
            var table = doc.AddTable(2, 2, WordTableStyle.TableGrid);

            var row0_first = table.Rows[0];
            var cell00 = row0_first.Cells[0];

            // Row -> Table
            Assert.True(row0_first.Parent == table);

            // Cell -> Row
            Assert.True(cell00.Parent == row0_first);

            // Cell -> Table
            Assert.True(cell00.ParentTable == table);

            // Paragraph parent back to cell, then up to row
            var para = cell00.Paragraphs.First();
            var parentCell = Assert.IsType<WordTableCell>(para.Parent);
            Assert.True(parentCell == cell00);
            Assert.True(parentCell.Parent == row0_first);
        }
    }
}

