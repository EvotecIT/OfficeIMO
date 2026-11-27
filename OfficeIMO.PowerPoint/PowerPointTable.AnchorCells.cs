using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint;

public partial class PowerPointTable {
    /// <summary>Enumerates visible anchors, including spans whose covered cells omit redundant merge flags.</summary>
    internal IEnumerable<(int Row, int Column, PowerPointTableCell Cell)> EnumerateAnchorCells() {
        List<List<OfficeXmlNode>> rows = CellNodes();
        int columns = ColumnCount;
        var coveredUntilRow = new int[columns];
        for (int row = 0; row < rows.Count; row++) {
            List<OfficeXmlNode> cells = rows[row];
            for (int column = 0; column < columns; column++) {
                if (coveredUntilRow[column] > row) continue;
                if (column >= cells.Count) throw new ArgumentOutOfRangeException(nameof(column), "The row has fewer cells than the table grid has columns.");
                PowerPointTableCell cell = CellFor(cells[column]);
                if (cell.IsMergedCell) continue;
                (int rowSpan, int columnSpan) = cell.Merge;
                int endColumn = column + Math.Min(columns - column, Math.Max(1, columnSpan));
                int endRow = row + Math.Min(rows.Count - row, Math.Max(1, rowSpan));
                for (int covered = column; covered < endColumn; covered++) coveredUntilRow[covered] = endRow;
                yield return (row, column, cell);
            }
        }
    }
}
