using System.Linq;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint;

public partial class PowerPointTable {
    /// <summary>Enumerates visible anchors, including spans whose covered cells omit redundant merge flags.</summary>
    internal IEnumerable<(int Row, int Column, PowerPointTableCell Cell)> EnumerateAnchorCells() {
        A.Table table = Frame.Graphic!.GraphicData!.GetFirstChild<A.Table>()!;
        A.TableRow[] rows = table.Elements<A.TableRow>().ToArray();
        int columns = table.TableGrid!.Elements<A.GridColumn>().Count();
        var coveredUntilRow = new int[columns];
        for (int row = 0; row < rows.Length; row++) {
            A.TableCell[] cells = rows[row].Elements<A.TableCell>().Take(columns).ToArray();
            for (int column = 0; column < columns; column++) {
                if (coveredUntilRow[column] > row) continue;
                PowerPointTableCell cell = WrapCell(cells[column]);
                if (cell.IsMergedCell) continue;
                (int rowSpan, int columnSpan) = cell.Merge;
                int endColumn = column + Math.Min(columns - column, Math.Max(1, columnSpan));
                int endRow = row + Math.Min(rows.Length - row, Math.Max(1, rowSpan));
                for (int covered = column; covered < endColumn; covered++) coveredUntilRow[covered] = endRow;
                yield return (row, column, cell);
            }
        }
    }
}
