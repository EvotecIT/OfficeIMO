using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>Resolves logical table-cell spans from physical Open XML rows and cells.</summary>
    internal static class WordTableCellSpanResolver {
        internal static int GetColumnSpan(WordTableCell cell) => GetColumnSpan(cell._tableCell);

        internal static int GetColumnSpan(TableCell cell) {
            int gridSpan = GetGridWidth(cell);
            if (gridSpan > 1) {
                return gridSpan;
            }

            if (cell.TableCellProperties?.HorizontalMerge?.Val?.Value != MergedCellValues.Restart) {
                return 1;
            }
            int span = 1;
            TableCell? continuation = cell.NextSibling<TableCell>();
            while (continuation != null) {
                HorizontalMerge? merge = continuation.TableCellProperties?.HorizontalMerge;
                if (merge == null || merge.Val?.Value == MergedCellValues.Restart) {
                    break;
                }
                span += GetGridWidth(continuation);
                continuation = continuation.NextSibling<TableCell>();
            }

            return span;
        }

        internal static int GetRowSpan(WordTableCell cell) {
            if (cell.VerticalMerge != WordCellMerge.Restart) {
                return 1;
            }

            List<WordTableRow> rows = cell.ParentTable.Rows;
            int rowIndex = rows.FindIndex(row => ReferenceEquals(row._tableRow, cell.Parent._tableRow));
            if (rowIndex < 0 || !TryGetLogicalColumn(cell._tableCell, out int logicalColumn)) {
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

        internal static bool TryGetLogicalColumn(TableCell target, out int logicalColumn) {
            TableRow? row = target.Parent as TableRow;
            logicalColumn = Math.Max(0, row?.TableRowProperties?.GetFirstChild<GridBefore>()?.Val?.Value ?? 0);
            if (row == null) return false;
            foreach (TableCell cell in row.Elements<TableCell>()) {
                if (ReferenceEquals(cell, target)) {
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

        private static int GetGridWidth(WordTableCell cell) => GetGridWidth(cell._tableCell);

        private static int GetGridWidth(TableCell cell) {
            int gridSpan = cell.TableCellProperties?.GetFirstChild<GridSpan>()?.Val?.Value ?? 1;
            return gridSpan > 0 ? gridSpan : 1;
        }
    }
}
