using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointTable {

        /// <summary>
        ///     Retrieves a cell at the specified row and column index.
        /// </summary>
        /// <param name="row">Zero-based row index.</param>
        /// <param name="column">Zero-based column index.</param>
        public PowerPointTableCell GetCell(int row, int column) {
            OfficeXmlNode tableRow = NthChild(TableNode, "tr", row);
            if (tableRow.IsNone) throw new ArgumentOutOfRangeException(nameof(row));
            OfficeXmlNode cell = NthChild(tableRow, "tc", column);
            if (cell.IsNone) throw new ArgumentOutOfRangeException(nameof(column));
            return CellFor(cell);
        }

        /// <summary>
        ///     Applies an action to all cells in the table.
        /// </summary>
        public void ApplyToCells(Action<PowerPointTableCell> action) {
            ArgumentNullException.ThrowIfNull(action);
            if (RowCount == 0 || ColumnCount == 0) {
                return;
            }

            ApplyToCells(0, RowCount - 1, 0, ColumnCount - 1, action);
        }

        /// <summary>
        ///     Applies an action to a rectangular range of cells.
        /// </summary>
        public void ApplyToCells(int startRow, int endRow, int startColumn, int endColumn, Action<PowerPointTableCell> action) {
            ArgumentNullException.ThrowIfNull(action);
            if (RowCount == 0 || ColumnCount == 0) {
                return;
            }

            int topRow = Math.Min(startRow, endRow);
            int bottomRow = Math.Max(startRow, endRow);
            int leftColumn = Math.Min(startColumn, endColumn);
            int rightColumn = Math.Max(startColumn, endColumn);

            if (topRow < 0 || leftColumn < 0) {
                throw new ArgumentOutOfRangeException("Row and column indices must be non-negative.");
            }
            if (bottomRow >= RowCount || rightColumn >= ColumnCount) {
                throw new ArgumentOutOfRangeException("Range exceeds table bounds.");
            }

            for (int r = topRow; r <= bottomRow; r++) {
                for (int c = leftColumn; c <= rightColumn; c++) {
                    action(GetCell(r, c));
                }
            }
        }

        /// <summary>
        ///     Applies an action to a specific row.
        /// </summary>
        public void ApplyToRow(int rowIndex, Action<PowerPointTableCell> action) {
            ArgumentOutOfRangeException.ThrowIfNegative(rowIndex);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(rowIndex, RowCount);
            ApplyToCells(rowIndex, rowIndex, 0, ColumnCount - 1, action);
        }

        /// <summary>
        ///     Applies an action to a specific column.
        /// </summary>
        public void ApplyToColumn(int columnIndex, Action<PowerPointTableCell> action) {
            ArgumentOutOfRangeException.ThrowIfNegative(columnIndex);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(columnIndex, ColumnCount);
            ApplyToCells(0, RowCount - 1, columnIndex, columnIndex, action);
        }

        /// <summary>
        ///     Sets cell padding for all cells.
        /// </summary>
        public PowerPointTable SetCellPadding(global::OfficeIMO.OfficeLength? left, global::OfficeIMO.OfficeLength? top, global::OfficeIMO.OfficeLength? right, global::OfficeIMO.OfficeLength? bottom) {
            ApplyToCells(cell => {
                cell.PaddingLeft = left;
                cell.PaddingTop = top;
                cell.PaddingRight = right;
                cell.PaddingBottom = bottom;
            });
            return this;
        }

        /// <summary>
        ///     Sets cell padding for a range of cells.
        /// </summary>
        public PowerPointTable SetCellPadding(global::OfficeIMO.OfficeLength? left, global::OfficeIMO.OfficeLength? top, global::OfficeIMO.OfficeLength? right, global::OfficeIMO.OfficeLength? bottom,
            int startRow, int endRow, int startColumn, int endColumn) {
            ApplyToCells(startRow, endRow, startColumn, endColumn, cell => {
                cell.PaddingLeft = left;
                cell.PaddingTop = top;
                cell.PaddingRight = right;
                cell.PaddingBottom = bottom;
            });
            return this;
        }

        /// <summary>
        ///     Sets cell alignment for all cells.
        /// </summary>
        public PowerPointTable SetCellAlignment(PowerPointTextAlignment? horizontal, PowerPointTextVerticalAlignment? vertical) {
            ApplyToCells(cell => {
                cell.HorizontalAlignment = horizontal;
                cell.VerticalAlignment = vertical;
            });
            return this;
        }

        /// <summary>
        ///     Sets cell alignment for a range of cells.
        /// </summary>
        public PowerPointTable SetCellAlignment(PowerPointTextAlignment? horizontal, PowerPointTextVerticalAlignment? vertical,
            int startRow, int endRow, int startColumn, int endColumn) {
            ApplyToCells(startRow, endRow, startColumn, endColumn, cell => {
                cell.HorizontalAlignment = horizontal;
                cell.VerticalAlignment = vertical;
            });
            return this;
        }

        /// <summary>Applies borders to all cells, retaining the existing width or dash when omitted.</summary>
        public void SetCellBorders(PowerPointTableCellBorders borders, OfficeColor color, OfficeLength? width = null,
            PowerPointLineDashStyle? dash = null) {
            PowerPointTableCell.ValidateBorderStyle(borders, color, width);
            if (dash.HasValue) _ = dash.Value.ToOpenXmlToken();
            ApplyToCells(cell => cell.SetBorders(borders, color, width, dash));
        }

        /// <summary>
        ///     Clears borders for all cells.
        /// </summary>
        public void ClearCellBorders(PowerPointTableCellBorders borders) {
            ApplyToCells(cell => cell.ClearBorders(borders));
        }

        /// <summary>
        ///     Merges a rectangular range of cells into the top-left cell.
        /// </summary>
        /// <param name="startRow">Zero-based start row.</param>
        /// <param name="startColumn">Zero-based start column.</param>
        /// <param name="endRow">Zero-based end row.</param>
        /// <param name="endColumn">Zero-based end column.</param>
        /// <param name="clearMergedContent">Whether to clear text from merged cells.</param>
        public void MergeCells(int startRow, int startColumn, int endRow, int endColumn, bool clearMergedContent = true) {
            int topRow = Math.Min(startRow, endRow);
            int bottomRow = Math.Max(startRow, endRow);
            int leftColumn = Math.Min(startColumn, endColumn);
            int rightColumn = Math.Max(startColumn, endColumn);

            if (topRow < 0 || leftColumn < 0) {
                throw new ArgumentOutOfRangeException("Row and column indices must be non-negative.");
            }
            if (bottomRow >= RowCount || rightColumn >= ColumnCount) {
                throw new ArgumentOutOfRangeException("Merge range exceeds table bounds.");
            }

            int rowSpan = bottomRow - topRow + 1;
            int colSpan = rightColumn - leftColumn + 1;
            if (rowSpan == 1 && colSpan == 1) {
                return;
            }

            DissolveMergesTouching(topRow, leftColumn, bottomRow, rightColumn);
            List<List<OfficeXmlNode>> cells = CellNodes();
            for (int r = topRow; r <= bottomRow; r++) {
                for (int c = leftColumn; c <= rightColumn; c++) {
                    OfficeXmlNode cell = cells[r][c];
                    if (r == topRow && c == leftColumn) {
                        SetMerge(cell, rowSpan > 1 ? rowSpan : null, colSpan > 1 ? colSpan : null, false, false);
                        continue;
                    }

                    SetMerge(cell, null, null, c > leftColumn, r > topRow);
                    if (clearMergedContent) {
                        ClearCellText(cell);
                    }
                }
            }
            MarkEdited();
        }

        // Sets a cell's spans and merge flags; null spans and false flags remove the attribute.
        private static void SetMerge(OfficeXmlNode cell, int? rowSpan, int? gridSpan, bool horizontal, bool vertical) {
            cell.SetInt32Attribute("rowSpan", rowSpan);
            cell.SetInt32Attribute("gridSpan", gridSpan);
            cell.SetAttribute("hMerge", horizontal ? "1" : null);
            cell.SetAttribute("vMerge", vertical ? "1" : null);
        }

        // Replaces the cell's paragraphs with one empty run, dropping the sounds the old text played.
        private void ClearCellText(OfficeXmlNode cell) {
            OfficeXmlNode body = cell.Child(D, "txBody");
            if (body.IsNone) return;
            List<string> soundIds = PowerPointTextBox.SoundIds(body);
            body.RemoveChildren(D, "p");
            OfficeXmlNode run = body.Append(body.CreateElement(D, "p")).Append(body.CreateElement(D, "r"));
            run.Append(run.CreateElement(D, "t"));
            RemoveUnusedSounds(soundIds);
        }

        private void RemoveUnusedSounds(List<string> soundIds) {
            if (soundIds.Count == 0) return;
            PowerPointPartRef owner = OwnerPartRef;
            if (owner.IsNone) return;
            foreach (string id in soundIds) owner.RemoveRelationshipIfUnused(id, default);
        }
    }
}
