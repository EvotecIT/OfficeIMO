namespace OfficeIMO.PowerPoint;

public partial class PowerPointTable {
    /// <summary>Enumerates visible anchors, including spans whose covered cells omit redundant merge flags.</summary>
    internal IEnumerable<(int Row, int Column, PowerPointTableCell Cell)> EnumerateAnchorCells() {
        var coveredUntilRow = new int[Columns];
        for (int row = 0; row < Rows; row++) {
            for (int column = 0; column < Columns; column++) {
                if (coveredUntilRow[column] > row) continue;
                PowerPointTableCell cell = GetCell(row, column);
                if (cell.IsMergedCell) continue;
                (int rowSpan, int columnSpan) = cell.Merge;
                int endColumn = column + Math.Min(Columns - column, Math.Max(1, columnSpan));
                int endRow = row + Math.Min(Rows - row, Math.Max(1, rowSpan));
                for (int covered = column; covered < endColumn; covered++) coveredUntilRow[covered] = endRow;
                yield return (row, column, cell);
            }
        }
    }
}
