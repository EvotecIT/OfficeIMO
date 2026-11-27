using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointTable {

        /// <summary>
        ///     Retrieves a row wrapper at the specified index.
        /// </summary>
        public PowerPointTableRow GetRow(int rowIndex) {
            ArgumentOutOfRangeException.ThrowIfNegative(rowIndex);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(rowIndex, RowCount);
            return new PowerPointTableRow(this, NthChild(TableNode, "tr", rowIndex));
        }

        /// <summary>
        ///     Retrieves a column wrapper at the specified index.
        /// </summary>
        public PowerPointTableColumn GetColumn(int columnIndex) {
            ArgumentOutOfRangeException.ThrowIfNegative(columnIndex);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(columnIndex, ColumnCount);
            return new PowerPointTableColumn(this, NthChild(GridNode, "gridCol", columnIndex));
        }

        // A new empty cell: a text body with one empty run, and empty cell properties.
        private static OfficeXmlNode CreateEmptyCell(OfficeXmlNode context) {
            OfficeXmlNode cell = context.CreateElement(D, "tc");
            OfficeXmlNode body = cell.Append(cell.CreateElement(D, "txBody"));
            body.Append(body.CreateElement(D, "bodyPr"));
            body.Append(body.CreateElement(D, "lstStyle"));
            OfficeXmlNode run = body.Append(body.CreateElement(D, "p")).Append(body.CreateElement(D, "r"));
            run.Append(run.CreateElement(D, "t"));
            cell.Append(cell.CreateElement(D, "tcPr"));
            return cell;
        }

        /// <summary>
        ///     Adds a new row to the table.
        /// </summary>
        /// <param name="index">Optional zero-based index where the row should be inserted. If omitted, row is appended.</param>
        public PowerPointTableRow AddRow(int? index = null) {
            if (index.HasValue) ArgumentOutOfRangeException.ThrowIfNegative(index.Value, nameof(index));
            OfficeXmlNode table = TableNode;
            int columns = ColumnCount;
            int rows = RowCount;
            var regions = CaptureMergeRegions();
            OfficeXmlNode row = table.CreateElement(D, "tr");
            row.SetInt64Attribute("h", 370840L);
            for (int c = 0; c < columns; c++) row.Append(CreateEmptyCell(row));

            int insertAt = index.HasValue && index.Value < rows ? index.Value : rows;
            row = table.InsertBefore(row, insertAt < rows ? NthChild(table, "tr", insertAt) : default);
            if (regions.Count > 0) {
                ShiftMergeRegionsForInsertedRow(regions, insertAt);
                ApplyMergeRegions(regions);
            }
            SyncFrameHeightToGrid();
            MarkEdited();
            return new PowerPointTableRow(this, row);
        }

        /// <summary>
        ///     Adds a new row cloned from a template row.
        /// </summary>
        public PowerPointTableRow AddRowFromTemplate(int templateRowIndex, int? index = null, bool clearText = true) {
            ArgumentOutOfRangeException.ThrowIfNegative(templateRowIndex);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(templateRowIndex, RowCount);
            if (index.HasValue) ArgumentOutOfRangeException.ThrowIfNegative(index.Value, nameof(index));

            OfficeXmlNode table = TableNode;
            int rows = RowCount;
            var regions = CaptureMergeRegions();
            var templateHorizontalRegions = regions
                .Where(region => region.Row <= templateRowIndex && templateRowIndex < region.Row + region.RowSpan
                    && region.ColumnSpan > 1)
                .Select(region => (region.Row, region.Column, RowSpan: 1, region.ColumnSpan))
                .ToList();
            OfficeXmlNode newRow = NthChild(table, "tr", templateRowIndex).Clone();
            if (clearText) {
                foreach (OfficeXmlNode cell in newRow.Children(D, "tc")) {
                    ClearCellText(cell);
                }
            }

            int insertAt = index.HasValue ? Math.Min(index.Value, rows) : rows;
            newRow = table.InsertBefore(newRow, insertAt < rows ? NthChild(table, "tr", insertAt) : default);
            if (regions.Count > 0) {
                ShiftMergeRegionsForInsertedRow(regions, insertAt);
                foreach (var region in templateHorizontalRegions) {
                    // Existing ranges expanded through the inserted row take precedence.
                    if (regions.Any(existing => existing.Row <= insertAt && insertAt < existing.Row + existing.RowSpan
                        && existing.Column < region.Column + region.ColumnSpan
                        && region.Column < existing.Column + existing.ColumnSpan)) continue;

                    regions.Add((insertAt, region.Column, region.RowSpan, region.ColumnSpan));
                }
                ApplyMergeRegions(regions);
            }

            SyncFrameHeightToGrid();
            MarkEdited();
            return new PowerPointTableRow(this, newRow);
        }

        /// <summary>
        ///     Removes a row at the specified index.
        /// </summary>
        /// <param name="index">Zero-based index of the row to remove.</param>
        public void RemoveRow(int index) {
            OfficeXmlNode row = NthChild(TableNode, "tr", index);
            if (row.IsNone) throw new ArgumentOutOfRangeException(nameof(index));
            RemoveRow(row);
        }

        internal void RemoveRow(OfficeXmlNode row) {
            if (row.Parent.IsNone) return;
            int index = TableNode.Children(D, "tr").ToList().IndexOf(row);
            var regions = index >= 0 ? CaptureMergeRegions() : new List<(int Row, int Column, int RowSpan, int ColumnSpan)>();
            PreserveAnchorsForRemovedRow(regions, index);
            List<string> soundIds = PowerPointTextBox.SoundIds(row);
            row.Remove();
            if (regions.Count > 0) {
                ShiftMergeRegionsForRemovedRow(regions, index);
                ApplyMergeRegions(regions);
            }
            RemoveUnusedSounds(soundIds);
            SyncFrameHeightToGrid();
            MarkEdited();
        }

        /// <summary>
        ///     Adds a new column to the table.
        /// </summary>
        /// <param name="index">Optional zero-based index where the column should be inserted. If omitted, column is appended.</param>
        public PowerPointTableColumn AddColumn(int? index = null) {
            if (index.HasValue) ArgumentOutOfRangeException.ThrowIfNegative(index.Value, nameof(index));
            OfficeXmlNode table = TableNode;
            OfficeXmlNode grid = GridNode;
            int existingColumns = ColumnCount;
            bool inside = index.HasValue && index.Value < existingColumns;
            var regions = CaptureMergeRegions();
            OfficeXmlNode gridColumn = grid.CreateElement(D, "gridCol");
            gridColumn.SetInt64Attribute("w", 3708400L);
            gridColumn = grid.InsertBefore(gridColumn, inside ? NthChild(grid, "gridCol", index!.Value) : default);

            foreach (OfficeXmlNode row in table.Children(D, "tr")) {
                OfficeXmlNode cell = CreateEmptyCell(row);
                row.InsertBefore(cell, inside ? NthChild(row, "tc", index!.Value) : row.Child(D, "extLst"));
            }
            if (regions.Count > 0) {
                ShiftMergeRegionsForInsertedColumn(regions, inside ? index!.Value : existingColumns);
                ApplyMergeRegions(regions);
            }
            SyncFrameWidthToGrid();
            MarkEdited();
            return new PowerPointTableColumn(this, gridColumn);
        }

        /// <summary>
        ///     Adds a new column cloned from a template column.
        /// </summary>
        public PowerPointTableColumn AddColumnFromTemplate(int templateColumnIndex, int? index = null, bool clearText = true) {
            ArgumentOutOfRangeException.ThrowIfNegative(templateColumnIndex);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(templateColumnIndex, ColumnCount);
            if (index.HasValue) ArgumentOutOfRangeException.ThrowIfNegative(index.Value, nameof(index));

            OfficeXmlNode table = TableNode;
            OfficeXmlNode grid = GridNode;
            if (grid.IsNone) throw new InvalidOperationException("Table grid is missing.");
            int existingColumns = ColumnCount;
            var regions = CaptureMergeRegions();
            var templateVerticalRegions = regions
                .Where(region => region.Column <= templateColumnIndex &&
                    templateColumnIndex < region.Column + region.ColumnSpan && region.RowSpan > 1)
                .Select(region => (region.Row, region.Column, region.RowSpan, ColumnSpan: 1))
                .ToList();
            OfficeXmlNode newColumn = NthChild(grid, "gridCol", templateColumnIndex).Clone();

            int insertAt = index.HasValue ? Math.Min(index.Value, existingColumns) : existingColumns;
            newColumn = grid.InsertBefore(newColumn, insertAt < existingColumns ? NthChild(grid, "gridCol", insertAt) : default);

            foreach (OfficeXmlNode row in table.Children(D, "tr")) {
                OfficeXmlNode newCell = NthChild(row, "tc", templateColumnIndex).Clone();
                if (clearText) {
                    ClearCellText(newCell);
                }
                row.InsertBefore(newCell, insertAt < existingColumns ? NthChild(row, "tc", insertAt) : row.Child(D, "extLst"));
            }
            if (regions.Count > 0) {
                ShiftMergeRegionsForInsertedColumn(regions, insertAt);
                foreach (var region in templateVerticalRegions) {
                    // Existing ranges expanded through the inserted column take precedence.
                    if (regions.Any(existing => existing.Column <= insertAt && insertAt < existing.Column + existing.ColumnSpan
                        && existing.Row < region.Row + region.RowSpan
                        && region.Row < existing.Row + existing.RowSpan)) continue;

                    regions.Add((region.Row, insertAt, region.RowSpan, region.ColumnSpan));
                }
                ApplyMergeRegions(regions);
            }

            SyncFrameWidthToGrid();
            MarkEdited();
            return new PowerPointTableColumn(this, newColumn);
        }

        /// <summary>
        ///     Removes a column at the specified index.
        /// </summary>
        /// <param name="index">Zero-based index of the column to remove.</param>
        public void RemoveColumn(int index) {
            OfficeXmlNode table = TableNode;
            OfficeXmlNode gridColumn = NthChild(GridNode, "gridCol", index);
            if (gridColumn.IsNone) throw new ArgumentOutOfRangeException(nameof(index));
            var regions = CaptureMergeRegions();
            PreserveAnchorsForRemovedColumn(regions, index);
            var discardedCells = new List<OfficeXmlNode>();
            foreach (OfficeXmlNode row in table.Children(D, "tr")) {
                OfficeXmlNode cell = NthChild(row, "tc", index);
                if (cell.IsNone) throw new ArgumentOutOfRangeException(nameof(index));
                discardedCells.Add(cell);
            }
            var soundIds = new List<string>();
            foreach (OfficeXmlNode cell in discardedCells) {
                foreach (string id in PowerPointTextBox.SoundIds(cell)) {
                    if (!soundIds.Contains(id)) soundIds.Add(id);
                }
            }
            gridColumn.Remove();
            foreach (OfficeXmlNode cell in discardedCells) {
                cell.Remove();
            }
            if (regions.Count > 0) {
                ShiftMergeRegionsForRemovedColumn(regions, index);
                ApplyMergeRegions(regions);
            }
            RemoveUnusedSounds(soundIds);
            SyncFrameWidthToGrid();
            MarkEdited();
        }
    }
}
