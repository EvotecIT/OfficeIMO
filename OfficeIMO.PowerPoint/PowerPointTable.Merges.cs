using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
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

        private static Int32Value? SpanOrNull(int span) => span > 1 ? new Int32Value(span) : null;

        private static void MoveAnchorContent(A.TableCell source, A.TableCell target) {
            if (source.TextBody != null) target.TextBody = (A.TextBody)source.TextBody.CloneNode(true);
            if (source.TableCellProperties != null) target.TableCellProperties = (A.TableCellProperties)source.TableCellProperties.CloneNode(true);
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

        /// <summary>A row inserted inside a vertical merge joins it, as it does when inserted in PowerPoint.</summary>
        private void ContinueMergesThroughRow(A.TableRow inserted, A.TableRow next) {
            A.TableCell[][] cells = CellMatrix();
            int nextIndex = TableElement.Elements<A.TableRow>().ToList().IndexOf(next);
            A.TableCell[] added = inserted.Elements<A.TableCell>().ToArray();
            for (int c = 0; c < added.Length && c < cells[nextIndex].Length; c++) {
                A.TableCell below = cells[nextIndex][c];
                if (!IsCoveredVertically(below)) continue;
                (int anchorRow, int anchorColumn) = FindMergeAnchor(cells, nextIndex, c);
                A.TableCell anchor = cells[anchorRow][anchorColumn];
                added[c].RowSpan = null;
                added[c].GridSpan = null;
                added[c].VerticalMerge = true;
                added[c].HorizontalMerge = IsCoveredHorizontally(below) ? true : (bool?)null;
                if (anchorColumn == c) anchor.RowSpan = (anchor.RowSpan?.Value ?? 1) + 1;
            }
        }

        /// <summary>A column inserted inside a horizontal merge joins it, as it does when inserted in PowerPoint.</summary>
        private void ContinueMergeThroughCell(A.TableCell inserted, A.TableCell next) {
            if (!IsCoveredHorizontally(next)) return;
            A.TableCell[][] cells = CellMatrix();
            int rowIndex = Array.FindIndex(cells, row => row.Contains(next));
            int column = Array.IndexOf(cells[rowIndex], next);
            int anchorColumn = column;
            while (anchorColumn > 0 && IsCoveredHorizontally(cells[rowIndex][anchorColumn])) anchorColumn--;
            A.TableCell left = cells[rowIndex][anchorColumn];
            inserted.RowSpan = null;
            inserted.GridSpan = null;
            inserted.HorizontalMerge = true;
            inserted.VerticalMerge = IsCoveredVertically(next) ? true : (bool?)null;
            if (!IsCoveredVertically(left)) left.GridSpan = (left.GridSpan?.Value ?? 1) + 1;
        }

        /// <summary>A copied row starts its own vertical extent, so only its horizontal merges survive.</summary>
        private static void KeepOnlyHorizontalMerges(A.TableRow row) {
            A.TableCell[] cells = row.Elements<A.TableCell>().ToArray();
            for (int c = 0; c < cells.Length; c++) {
                A.TableCell cell = cells[c];
                if (IsCoveredVertically(cell) && !IsCoveredHorizontally(cell)) {
                    int span = 1;
                    while (c + span < cells.Length && IsCoveredHorizontally(cells[c + span])) span++;
                    cell.GridSpan = SpanOrNull(span);
                }
                cell.RowSpan = null;
                cell.VerticalMerge = null;
            }
        }

        /// <summary>A copied column starts its own horizontal extent, so only its vertical merges survive.</summary>
        private static void KeepOnlyVerticalMerge(A.TableCell cell) {
            if (IsCoveredHorizontally(cell)) {
                cell.VerticalMerge = null;
                cell.RowSpan = null;
            }
            cell.GridSpan = null;
            cell.HorizontalMerge = null;
        }

        /// <summary>Shrinks merges crossing a row about to be removed; an anchor hands its content to the next row.</summary>
        private void ReleaseMergesInRow(A.TableRow row) {
            A.TableCell[][] cells = CellMatrix();
            int rowIndex = TableElement.Elements<A.TableRow>().ToList().IndexOf(row);
            if (rowIndex < 0) return;
            for (int c = 0; c < cells[rowIndex].Length; c++) {
                A.TableCell cell = cells[rowIndex][c];
                if (IsCoveredVertically(cell)) {
                    (int anchorRow, int anchorColumn) = FindMergeAnchor(cells, rowIndex, c);
                    if (anchorColumn == c) {
                        A.TableCell anchor = cells[anchorRow][anchorColumn];
                        anchor.RowSpan = SpanOrNull((anchor.RowSpan?.Value ?? 1) - 1);
                    }
                } else if (rowIndex + 1 < cells.Length && IsCoveredVertically(cells[rowIndex + 1][c])) {
                    A.TableCell below = cells[rowIndex + 1][c];
                    Int32Value? rowSpan = SpanOrNull((cell.RowSpan?.Value ?? 1) - 1);
                    Int32Value? gridSpan = SpanOrNull(cell.GridSpan?.Value ?? 1);
                    MoveAnchorContent(cell, below);
                    below.VerticalMerge = null;
                    below.RowSpan = rowSpan;
                    below.GridSpan = gridSpan;
                }
            }
        }

        /// <summary>Shrinks merges crossing a column about to be removed; an anchor hands its content to the next column.</summary>
        private void ReleaseMergesInColumn(int column) {
            A.TableCell[][] cells = CellMatrix();
            for (int r = 0; r < cells.Length; r++) {
                A.TableCell cell = cells[r][column];
                if (IsCoveredHorizontally(cell)) {
                    int anchorColumn = column;
                    while (anchorColumn > 0 && IsCoveredHorizontally(cells[r][anchorColumn])) anchorColumn--;
                    A.TableCell anchor = cells[r][anchorColumn];
                    anchor.GridSpan = SpanOrNull((anchor.GridSpan?.Value ?? 1) - 1);
                } else if (column + 1 < cells[r].Length && IsCoveredHorizontally(cells[r][column + 1])) {
                    A.TableCell right = cells[r][column + 1];
                    Int32Value? gridSpan = SpanOrNull((cell.GridSpan?.Value ?? 1) - 1);
                    Int32Value? rowSpan = SpanOrNull(cell.RowSpan?.Value ?? 1);
                    MoveAnchorContent(cell, right);
                    right.HorizontalMerge = null;
                    right.VerticalMerge = IsCoveredVertically(cell) ? true : (bool?)null;
                    right.GridSpan = gridSpan;
                    right.RowSpan = rowSpan;
                }
            }
        }
    }
}
