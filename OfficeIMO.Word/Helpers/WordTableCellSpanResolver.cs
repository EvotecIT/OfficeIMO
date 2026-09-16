using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>Resolves logical table-cell spans from physical Open XML rows and cells.</summary>
    internal static class WordTableCellSpanResolver {
        internal static int GetColumnSpan(WordTableCell cell) {
            int gridSpan = GetGridWidth(cell);
            if (gridSpan > 1) {
                return gridSpan;
            }

            if (cell.HorizontalMerge != WordCellMerge.Restart) {
                return 1;
            }

            List<WordTableCell> cells = cell.Parent.Cells;
            int columnIndex = FindCellIndex(cells, cell);
            if (columnIndex < 0) {
                return 1;
            }

            int span = 1;
            for (int index = columnIndex + 1; index < cells.Count; index++) {
                if (cells[index].HorizontalMerge != WordCellMerge.Continue) {
                    break;
                }

                span += GetGridWidth(cells[index]);
            }

            return span;
        }

        internal static int GetRowSpan(WordTableCell cell) {
            if (cell.VerticalMerge != WordCellMerge.Restart) {
                return 1;
            }

            List<WordTableRow> rows = cell.ParentTable.Rows;
            int rowIndex = rows.FindIndex(row => ReferenceEquals(row._tableRow, cell.Parent._tableRow));
            if (rowIndex < 0 || !TryGetLogicalColumn(cell.Parent, cell, out int logicalColumn)) {
                return 1;
            }

            int logicalWidth = GetColumnSpan(cell);
            int span = 1;
            for (int index = rowIndex + 1; index < rows.Count; index++) {
                WordTableCell? continuation = FindCellStartingAtLogicalColumn(rows[index], logicalColumn);
                if (continuation?.VerticalMerge != WordCellMerge.Continue
                    || GetColumnSpan(continuation) != logicalWidth) {
                    break;
                }

                span++;
            }

            return span;
        }

        private static bool TryGetLogicalColumn(WordTableRow row, WordTableCell target, out int logicalColumn) {
            logicalColumn = GetGridBefore(row);
            foreach (WordTableCell cell in row.Cells) {
                if (ReferenceEquals(cell._tableCell, target._tableCell)) {
                    return true;
                }

                logicalColumn += GetGridWidth(cell);
            }

            return false;
        }

        private static WordTableCell? FindCellStartingAtLogicalColumn(WordTableRow row, int logicalColumn) {
            int currentColumn = GetGridBefore(row);
            foreach (WordTableCell cell in row.Cells) {
                if (currentColumn == logicalColumn) {
                    return cell;
                }

                if (currentColumn > logicalColumn) {
                    return null;
                }

                currentColumn += GetGridWidth(cell);
            }

            return null;
        }

        private static int GetGridBefore(WordTableRow row) {
            int gridBefore = row._tableRow.TableRowProperties?.GetFirstChild<GridBefore>()?.Val?.Value ?? 0;
            return gridBefore > 0 ? gridBefore : 0;
        }

        private static int GetGridWidth(WordTableCell cell) {
            int gridSpan = cell._tableCell.TableCellProperties?.GetFirstChild<GridSpan>()?.Val?.Value ?? 1;
            return gridSpan > 0 ? gridSpan : 1;
        }

        private static int FindCellIndex(IReadOnlyList<WordTableCell> cells, WordTableCell target) {
            for (int index = 0; index < cells.Count; index++) {
                if (ReferenceEquals(cells[index]._tableCell, target._tableCell)) {
                    return index;
                }
            }

            return -1;
        }
    }
}
