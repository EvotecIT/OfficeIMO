using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointTable {

        /// <summary>
        ///     Retrieves a row wrapper at the specified index.
        /// </summary>
        public PowerPointTableRow GetRow(int rowIndex) {
            ArgumentOutOfRangeException.ThrowIfNegative(rowIndex);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(rowIndex, Rows);

            A.TableRow tableRow = TableElement.Elements<A.TableRow>().ElementAt(rowIndex);
            return new PowerPointTableRow(this, tableRow);
        }

        /// <summary>
        ///     Retrieves a column wrapper at the specified index.
        /// </summary>
        public PowerPointTableColumn GetColumn(int columnIndex) {
            ArgumentOutOfRangeException.ThrowIfNegative(columnIndex);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(columnIndex, Columns);

            A.GridColumn column = TableElement.TableGrid!.Elements<A.GridColumn>().ElementAt(columnIndex);
            return new PowerPointTableColumn(this, column);
        }

        /// <summary>
        ///     Adds a new row to the table.
        /// </summary>
        /// <param name="index">Optional zero-based index where the row should be inserted. If omitted, row is appended.</param>
        public PowerPointTableRow AddRow(int? index = null) {
            A.Table table = Frame.Graphic!.GraphicData!.GetFirstChild<A.Table>()!;
            int columns = Columns;
            var regions = CaptureMergeRegions();
            A.TableRow row = new() { Height = 370840L };
            for (int c = 0; c < columns; c++) {
                A.TableCell cell = new(
                    new A.TextBody(new A.BodyProperties(), new A.ListStyle(),
                        new A.Paragraph(new A.Run(new A.Text(string.Empty)))),
                    new A.TableCellProperties()
                );
                row.Append(cell);
            }

            int insertAt = index.HasValue && index.Value < Rows ? index.Value : Rows;
            if (insertAt < Rows) {
                A.TableRow refRow = table.Elements<A.TableRow>().ElementAt(insertAt);
                table.InsertBefore(row, refRow);
            } else {
                table.Append(row);
            }
            if (regions.Count > 0) {
                ShiftMergeRegionsForInsertedRow(regions, insertAt);
                ApplyMergeRegions(regions);
            }
            SyncFrameHeightToGrid();
            return new PowerPointTableRow(this, row);
        }

        /// <summary>
        ///     Adds a new row cloned from a template row.
        /// </summary>
        public PowerPointTableRow AddRowFromTemplate(int templateRowIndex, int? index = null, bool clearText = true) {
            ArgumentOutOfRangeException.ThrowIfNegative(templateRowIndex);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(templateRowIndex, Rows);

            A.Table table = TableElement;
            var regions = CaptureMergeRegions();
            var templateHorizontalRegions = regions
                .Where(region => region.Row <= templateRowIndex && templateRowIndex < region.Row + region.RowSpan
                    && region.ColumnSpan > 1)
                .Select(region => (region.Row, region.Column, RowSpan: 1, region.ColumnSpan))
                .ToList();
            A.TableRow templateRow = table.Elements<A.TableRow>().ElementAt(templateRowIndex);
            A.TableRow newRow = (A.TableRow)templateRow.CloneNode(true);
            if (clearText) {
                foreach (A.TableCell cell in newRow.Elements<A.TableCell>()) {
                    ClearCellText(cell);
                }
            }

            int insertAt = index.HasValue ? Math.Min(index.Value, Rows) : Rows;
            if (insertAt < Rows) {
                A.TableRow refRow = table.Elements<A.TableRow>().ElementAt(insertAt);
                table.InsertBefore(newRow, refRow);
            } else {
                table.Append(newRow);
            }
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
            return new PowerPointTableRow(this, newRow);
        }

        /// <summary>
        ///     Removes a row at the specified index.
        /// </summary>
        /// <param name="index">Zero-based index of the row to remove.</param>
        public void RemoveRow(int index) {
            A.Table table = Frame.Graphic!.GraphicData!.GetFirstChild<A.Table>()!;
            A.TableRow row = table.Elements<A.TableRow>().ElementAt(index);
            RemoveRow(row);
        }

        internal void RemoveRow(A.TableRow row) {
            if (row.Parent == null) return;
            int index = TableElement.Elements<A.TableRow>().ToList().IndexOf(row);
            var regions = index >= 0 ? CaptureMergeRegions() : new List<(int Row, int Column, int RowSpan, int ColumnSpan)>();
            PreserveAnchorsForRemovedRow(regions, index);
            string[] discardedSoundIds = PowerPointEmbeddedSound
                .GetRelationshipIds(row);
            row.Remove();
            if (regions.Count > 0) {
                ShiftMergeRegionsForRemovedRow(regions, index);
                ApplyMergeRegions(regions);
            }
            PowerPointEmbeddedSound.RemoveIfUnused(RelationshipOwnerPart ?? SlidePart,
                discardedSoundIds);
            SyncFrameHeightToGrid();
        }

        /// <summary>
        ///     Adds a new column to the table.
        /// </summary>
        /// <param name="index">Optional zero-based index where the column should be inserted. If omitted, column is appended.</param>
        public PowerPointTableColumn AddColumn(int? index = null) {
            A.Table table = Frame.Graphic!.GraphicData!.GetFirstChild<A.Table>()!;
            A.TableGrid grid = table.TableGrid!;
            int existingColumns = Columns;
            var regions = CaptureMergeRegions();
            A.GridColumn gridColumn = new() { Width = 3708400L };

            if (index.HasValue && index.Value < existingColumns) {
                A.GridColumn refCol = grid.Elements<A.GridColumn>().ElementAt(index.Value);
                grid.InsertBefore(gridColumn, refCol);
            } else {
                grid.Append(gridColumn);
            }

            foreach (A.TableRow row in table.Elements<A.TableRow>()) {
                A.TableCell cell = new(
                    new A.TextBody(new A.BodyProperties(), new A.ListStyle(),
                        new A.Paragraph(new A.Run(new A.Text(string.Empty)))),
                    new A.TableCellProperties()
                );

                if (index.HasValue && index.Value < existingColumns) {
                    A.TableCell refCell = row.Elements<A.TableCell>().ElementAt(index.Value);
                    row.InsertBefore(cell, refCell);
                } else {
                    row.InsertBefore(cell, row.GetFirstChild<A.ExtensionList>());
                }
            }
            if (regions.Count > 0) {
                ShiftMergeRegionsForInsertedColumn(regions,
                    index.HasValue && index.Value < existingColumns ? index.Value : existingColumns);
                ApplyMergeRegions(regions);
            }
            SyncFrameWidthToGrid();
            return new PowerPointTableColumn(this, gridColumn);
        }

        /// <summary>
        ///     Adds a new column cloned from a template column.
        /// </summary>
        public PowerPointTableColumn AddColumnFromTemplate(int templateColumnIndex, int? index = null, bool clearText = true) {
            ArgumentOutOfRangeException.ThrowIfNegative(templateColumnIndex);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(templateColumnIndex, Columns);

            A.Table table = TableElement;
            A.TableGrid grid = table.TableGrid ?? throw new InvalidOperationException("Table grid is missing.");
            int existingColumns = Columns;
            var regions = CaptureMergeRegions();
            var templateVerticalRegions = regions
                .Where(region => region.Column <= templateColumnIndex &&
                    templateColumnIndex < region.Column + region.ColumnSpan && region.RowSpan > 1)
                .Select(region => (region.Row, region.Column, region.RowSpan, ColumnSpan: 1))
                .ToList();
            A.GridColumn templateColumn = grid.Elements<A.GridColumn>().ElementAt(templateColumnIndex);
            A.GridColumn newColumn = (A.GridColumn)templateColumn.CloneNode(true);

            int insertAt = index.HasValue ? Math.Min(index.Value, existingColumns) : existingColumns;
            if (insertAt < existingColumns) {
                A.GridColumn refColumn = grid.Elements<A.GridColumn>().ElementAt(insertAt);
                grid.InsertBefore(newColumn, refColumn);
            } else {
                grid.Append(newColumn);
            }

            foreach (A.TableRow row in table.Elements<A.TableRow>()) {
                A.TableCell templateCell = row.Elements<A.TableCell>().ElementAt(templateColumnIndex);
                A.TableCell newCell = (A.TableCell)templateCell.CloneNode(true);
                if (clearText) {
                    ClearCellText(newCell);
                }

                if (insertAt < existingColumns) {
                    A.TableCell refCell = row.Elements<A.TableCell>().ElementAt(insertAt);
                    row.InsertBefore(newCell, refCell);
                } else {
                    row.InsertBefore(newCell, row.GetFirstChild<A.ExtensionList>());
                }
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
            return new PowerPointTableColumn(this, newColumn);
        }

        /// <summary>
        ///     Removes a column at the specified index.
        /// </summary>
        /// <param name="index">Zero-based index of the column to remove.</param>
        public void RemoveColumn(int index) {
            A.Table table = Frame.Graphic!.GraphicData!.GetFirstChild<A.Table>()!;
            A.TableGrid grid = table.TableGrid!;
            var regions = CaptureMergeRegions();
            PreserveAnchorsForRemovedColumn(regions, index);
            A.TableCell[] discardedCells = table.Elements<A.TableRow>()
                .Select(row => row.Elements<A.TableCell>().ElementAt(index))
                .ToArray();
            string[] discardedSoundIds = discardedCells
                .SelectMany(PowerPointEmbeddedSound.GetRelationshipIds)
                .Distinct(StringComparer.Ordinal)
                .ToArray();
            grid.Elements<A.GridColumn>().ElementAt(index).Remove();
            foreach (A.TableCell cell in discardedCells) {
                cell.Remove();
            }
            if (regions.Count > 0) {
                ShiftMergeRegionsForRemovedColumn(regions, index);
                ApplyMergeRegions(regions);
            }
            PowerPointEmbeddedSound.RemoveIfUnused(RelationshipOwnerPart ?? SlidePart,
                discardedSoundIds);
            SyncFrameWidthToGrid();
        }
    }
}
