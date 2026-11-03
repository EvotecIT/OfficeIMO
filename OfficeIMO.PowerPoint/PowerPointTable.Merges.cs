using System;
using System.Collections.Generic;
using System.Linq;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointTable {
        // Every grid position holds one a:tc: an anchor carries rowSpan/gridSpan and the cells it covers carry
        // vMerge/hMerge. PowerPoint repairs tables whose spans and covered cells disagree, so structural edits
        // keep both sides in step.

        private A.TableCell[][] CellMatrix() => TableElement.Elements<A.TableRow>()
            .Select(row => row.Elements<A.TableCell>().ToArray()).ToArray();

        private static bool IsCoveredHorizontally(A.TableCell cell) => cell.HorizontalMerge?.Value == true;
        private static bool IsCoveredVertically(A.TableCell cell) => cell.VerticalMerge?.Value == true;

        private static (int Row, int Column) FindMergeAnchor(A.TableCell[][] cells, int row, int column) {
            while (column > 0 && IsCoveredHorizontally(cells[row][column])) column--;
            while (row > 0 && IsCoveredVertically(cells[row][column])) row--;
            return (row, column);
        }

        private void DissolveMergesTouching(int topRow, int leftColumn, int bottomRow, int rightColumn) {
            A.TableCell[][] cells = CellMatrix();
            var anchors = new HashSet<(int Row, int Column)>();
            for (int r = topRow; r <= bottomRow; r++) {
                for (int c = leftColumn; c <= rightColumn; c++) anchors.Add(FindMergeAnchor(cells, r, c));
            }

            foreach ((int anchorRow, int anchorColumn) in anchors) {
                A.TableCell anchor = cells[anchorRow][anchorColumn];
                int lastRow = Math.Min(cells.Length, anchorRow + (anchor.RowSpan?.Value ?? 1));
                int lastColumn = Math.Min(cells[anchorRow].Length, anchorColumn + (anchor.GridSpan?.Value ?? 1));
                for (int r = anchorRow; r < lastRow; r++) {
                    for (int c = anchorColumn; c < lastColumn; c++) {
                        A.TableCell cell = cells[r][c];
                        cell.RowSpan = null;
                        cell.GridSpan = null;
                        cell.HorizontalMerge = null;
                        cell.VerticalMerge = null;
                    }
                }
            }
        }

    }
}
