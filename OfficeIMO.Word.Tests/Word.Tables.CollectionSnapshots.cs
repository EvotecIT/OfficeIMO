using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void TableCollections_ReturnFreshLiveSnapshotsAndReuseElementWrappers() {
            using WordDocument document = WordDocument.Create();
            WordTable table = document.AddTable(2, 2);

            List<WordTableRow> firstRows = table.Rows;
            List<WordTableRow> secondRows = table.Rows;

            Assert.NotSame(firstRows, secondRows);
            Assert.Same(firstRows[0], secondRows[0]);
            firstRows.RemoveAt(0);
            Assert.Equal(2, table.Rows.Count);

            List<WordTableCell> firstCells = table.Rows[0].Cells;
            List<WordTableCell> secondCells = table.Rows[0].Cells;

            Assert.NotSame(firstCells, secondCells);
            Assert.Same(firstCells[0], secondCells[0]);
            firstCells.Clear();
            Assert.Equal(2, table.Rows[0].Cells.Count);

            WordTableCell firstCell = secondCells[0];
            WordTableCell adjacentCell = secondCells[1];
            WordTableCell cellBelow = table.Rows[1].Cells[0];
            firstCell.MergeHorizontally(1);
            Assert.Same(firstCell, table.Rows[0].Cells[0]);
            Assert.Same(adjacentCell, table.Rows[0].Cells[1]);
            firstCell.SplitHorizontally(1);
            firstCell.MergeVertically(1);
            Assert.Same(firstCell, table.Rows[0].Cells[0]);
            Assert.Same(cellBelow, table.Rows[1].Cells[0]);
            firstCell.SplitVertically(1);

            WordTableRow addedRow = table.AddRow(2);
            Assert.Same(addedRow, table.Rows[2]);

            WordTableCell removedCell = addedRow.Cells[0];
            removedCell.Remove();
            Assert.Single(addedRow.Cells);

            addedRow.Remove();
            Assert.Equal(2, table.Rows.Count);
        }
    }
}
