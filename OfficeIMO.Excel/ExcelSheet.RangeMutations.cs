using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>Plans insertion of a rectangular cell block.</summary>
        public ExcelStructuralMutationPlan PlanInsertCells(
            string range,
            ExcelCellShiftDirection direction,
            ExcelMutationPlanOptions? options = null) {
            if (direction != ExcelCellShiftDirection.Right && direction != ExcelCellShiftDirection.Down) {
                throw new ArgumentOutOfRangeException(nameof(direction), "Cell insertion shifts right or down.");
            }
            return PlanCellShift(range, direction, inserting: true, options);
        }

        /// <summary>Plans deletion of a rectangular cell block.</summary>
        public ExcelStructuralMutationPlan PlanDeleteCells(
            string range,
            ExcelCellShiftDirection direction,
            ExcelMutationPlanOptions? options = null) {
            if (direction != ExcelCellShiftDirection.Left && direction != ExcelCellShiftDirection.Up) {
                throw new ArgumentOutOfRangeException(nameof(direction), "Cell deletion shifts left or up.");
            }
            return PlanCellShift(range, direction, inserting: false, options);
        }

        /// <summary>Transactionally inserts cells.</summary>
        public ExcelMutationResult InsertCells(string range, ExcelCellShiftDirection direction, ExcelMutationPlanOptions? options = null, CancellationToken cancellationToken = default) =>
            PlanInsertCells(range, direction, options).Apply(cancellationToken);

        /// <summary>Transactionally deletes cells.</summary>
        public ExcelMutationResult DeleteCells(string range, ExcelCellShiftDirection direction, ExcelMutationPlanOptions? options = null, CancellationToken cancellationToken = default) =>
            PlanDeleteCells(range, direction, options).Apply(cancellationToken);

        /// <summary>Plans a formula-aware cell-range copy.</summary>
        public ExcelStructuralMutationPlan PlanCopyRange(string sourceRange, string destinationTopLeft, ExcelMutationPlanOptions? options = null) =>
            PlanRangeTransfer(sourceRange, destinationTopLeft, move: false, transpose: false, options);

        /// <summary>Plans a formula-aware cell-range move.</summary>
        public ExcelStructuralMutationPlan PlanMoveRange(string sourceRange, string destinationTopLeft, ExcelMutationPlanOptions? options = null) =>
            PlanRangeTransfer(sourceRange, destinationTopLeft, move: true, transpose: false, options);

        /// <summary>Plans a formula-aware transposed copy.</summary>
        public ExcelStructuralMutationPlan PlanTransposeRange(string sourceRange, string destinationTopLeft, ExcelMutationPlanOptions? options = null) =>
            PlanRangeTransfer(sourceRange, destinationTopLeft, move: false, transpose: true, options);

        /// <summary>Copies cells, styles, formulas, and cell-anchored images.</summary>
        public ExcelMutationResult CopyRange(string sourceRange, string destinationTopLeft, ExcelMutationPlanOptions? options = null, CancellationToken cancellationToken = default) =>
            PlanCopyRange(sourceRange, destinationTopLeft, options).Apply(cancellationToken);

        /// <summary>Moves cells, styles, formulas, and cell-anchored images.</summary>
        public ExcelMutationResult MoveRange(string sourceRange, string destinationTopLeft, ExcelMutationPlanOptions? options = null, CancellationToken cancellationToken = default) =>
            PlanMoveRange(sourceRange, destinationTopLeft, options).Apply(cancellationToken);

        /// <summary>Copies a range with rows and columns transposed.</summary>
        public ExcelMutationResult TransposeRange(string sourceRange, string destinationTopLeft, ExcelMutationPlanOptions? options = null, CancellationToken cancellationToken = default) =>
            PlanTransposeRange(sourceRange, destinationTopLeft, options).Apply(cancellationToken);

        private ExcelStructuralMutationPlan PlanCellShift(
            string range,
            ExcelCellShiftDirection direction,
            bool inserting,
            ExcelMutationPlanOptions? options) {
            ExcelReference affected = ParseLocalCellRange(range);
            ExcelMutationPlanOptions effective = (options ?? new ExcelMutationPlanOptions()).CloneAndValidate();
            affected.GetBounds(out int r1, out int c1, out int r2, out int c2);
            return Locking.ExecuteRead(_excelDocument.EnsureLock(), () => {
                EnsureMutationPlanCanInspectWithoutMaterializing();
                PreflightCellShift(affected, direction, inserting);
                int count = WorksheetRoot.Descendants<Cell>().Count(cell => {
                    if (!TryGetCellCoordinates(cell, out int row, out int column)) return false;
                    return direction == ExcelCellShiftDirection.Right || direction == ExcelCellShiftDirection.Left
                        ? row >= r1 && row <= r2 && column >= c1
                        : column >= c1 && column <= c2 && row >= r1;
                });
                if (count > effective.MaximumAffectedCells) throw new InvalidOperationException($"Cell shift affects {count} cells, exceeding MaximumAffectedCells ({effective.MaximumAffectedCells}).");
                var impacts = new[] { new ExcelMutationImpact("cells", count, "Cells in the selected row or column band will be shifted.") };
                ExcelStructuralMutationKind kind = inserting
                    ? direction == ExcelCellShiftDirection.Right ? ExcelStructuralMutationKind.InsertCellsRight : ExcelStructuralMutationKind.InsertCellsDown
                    : direction == ExcelCellShiftDirection.Left ? ExcelStructuralMutationKind.DeleteCellsLeft : ExcelStructuralMutationKind.DeleteCellsUp;
                return new ExcelStructuralMutationPlan(this, kind, affected.ToString(), null, count, impacts, effective,
                    cancellationToken => {
                        cancellationToken.ThrowIfCancellationRequested();
                        _ = PlanCellShift(affected.ToString(), direction, inserting, effective);
                        ApplyCellShift(affected, direction, inserting, cancellationToken);
                    });
            });
        }

        private ExcelStructuralMutationPlan PlanRangeTransfer(
            string sourceRange,
            string destinationTopLeft,
            bool move,
            bool transpose,
            ExcelMutationPlanOptions? options) {
            ExcelReference source = ParseLocalCellRange(sourceRange);
            ExcelReference destination = ExcelReference.Parse(destinationTopLeft);
            if (destination.Kind != ExcelReferenceKind.Cell || destination.IsQualified) throw new ArgumentException("Destination must be one local A1 cell.", nameof(destinationTopLeft));
            source.GetBounds(out int r1, out int c1, out int r2, out int c2);
            int rows = r2 - r1 + 1;
            int columns = c2 - c1 + 1;
            int destinationRows = transpose ? columns : rows;
            int destinationColumns = transpose ? rows : columns;
            if ((long)destination.Start.Row + destinationRows - 1L > A1.MaxRows
                || (long)destination.Start.Column + destinationColumns - 1L > A1.MaxColumns) {
                throw new ArgumentOutOfRangeException(nameof(destinationTopLeft), "Destination range exceeds worksheet limits.");
            }
            long area = (long)rows * columns;
            ExcelMutationPlanOptions effective = (options ?? new ExcelMutationPlanOptions()).CloneAndValidate();
            if (area > effective.MaximumAffectedCells) throw new InvalidOperationException($"Range transfer affects {area} cells, exceeding MaximumAffectedCells ({effective.MaximumAffectedCells}).");
            return Locking.ExecuteRead(_excelDocument.EnsureLock(), () => {
                EnsureMutationPlanCanInspectWithoutMaterializing();
                PreflightRangeTransfer(source, destination.Start.Row, destination.Start.Column, destinationRows, destinationColumns, move, transpose);
                int existing = WorksheetRoot.Descendants<Cell>().Count(cell =>
                    TryGetCellCoordinates(cell, out int row, out int column)
                    && source.Contains(row, column));
                int images = Images.Count(image => source.Contains(image.RowIndex, image.ColumnIndex));
                var impacts = new List<ExcelMutationImpact> {
                    new ExcelMutationImpact("cells", existing, "Existing cells, formulas, and styles will be transferred.")
                };
                if (images > 0) impacts.Add(new ExcelMutationImpact("images", images, "Cell-anchored images will follow the transfer."));
                ExcelStructuralMutationKind kind = transpose ? ExcelStructuralMutationKind.Transpose : move ? ExcelStructuralMutationKind.Move : ExcelStructuralMutationKind.Copy;
                return new ExcelStructuralMutationPlan(this, kind, source.ToString(), destination.ToString(), (int)area, impacts, effective,
                    cancellationToken => {
                        cancellationToken.ThrowIfCancellationRequested();
                        _ = PlanRangeTransfer(source.ToString(), destination.ToString(), move, transpose, effective);
                        ApplyRangeTransfer(source, destination.Start.Row, destination.Start.Column, move, transpose, cancellationToken);
                    });
            });
        }

        private void PreflightCellShift(ExcelReference affected, ExcelCellShiftDirection direction, bool inserting) {
            EnsureNoIntersectingOwnedStructures(affected, "Cell shifts cannot split tables, merged cells, array formulas, data tables, or PivotTable output ranges.");
            affected.GetBounds(out int r1, out int c1, out int r2, out int c2);
            if (!inserting) return;
            if (direction == ExcelCellShiftDirection.Right) {
                int max = WorksheetRoot.Descendants<Cell>().Where(cell => TryGetCellCoordinates(cell, out int row, out _) && row >= r1 && row <= r2)
                    .Select(cell => TryGetCellCoordinates(cell, out _, out int column) ? column : 0).DefaultIfEmpty().Max();
                if ((long)max + c2 - c1 + 1L > A1.MaxColumns) throw new InvalidOperationException("Cell insertion would exceed the worksheet column limit.");
            } else {
                int max = WorksheetRoot.Descendants<Cell>().Where(cell => TryGetCellCoordinates(cell, out _, out int column) && column >= c1 && column <= c2)
                    .Select(cell => TryGetCellCoordinates(cell, out int row, out _) ? row : 0).DefaultIfEmpty().Max();
                if ((long)max + r2 - r1 + 1L > A1.MaxRows) throw new InvalidOperationException("Cell insertion would exceed the worksheet row limit.");
            }
        }

        private void PreflightRangeTransfer(ExcelReference source, int destinationRow, int destinationColumn, int rows, int columns, bool move, bool transpose) {
            ExcelReference destination = ExcelReference.Parse(A1.CellReference(destinationRow, destinationColumn) + ":" + A1.CellReference(destinationRow + rows - 1, destinationColumn + columns - 1));
            if (move || transpose) {
                EnsureNoIntersectingOwnedStructures(source, "Move and transpose cannot split owned table, merge, array, data-table, or PivotTable structures.");
            }
            EnsureNoIntersectingOwnedStructures(destination, "Range-transfer destination cannot overwrite owned table, merge, array, data-table, or PivotTable structures.");
            foreach (Cell cell in WorksheetRoot.Descendants<Cell>().Where(cell => TryGetCellCoordinates(cell, out int row, out int column) && source.Contains(row, column))) {
                CellFormulaValues? type = cell.CellFormula?.FormulaType?.Value;
                if (type == CellFormulaValues.Shared || type == CellFormulaValues.Array || type == CellFormulaValues.DataTable) {
                    throw new InvalidOperationException("Range transfer requires materialized ordinary formulas; shared, array, and data-table formulas cannot be split.");
                }
            }
        }

        private void EnsureNoIntersectingOwnedStructures(ExcelReference range, string message) {
            bool conflict = _worksheetPart.TableDefinitionParts.Any(part => ExcelReference.TryParse(part.Table?.Reference?.Value, out ExcelReference? table) && table!.Intersects(range))
                || WorksheetRoot.Descendants<MergeCell>().Any(merge => ExcelReference.TryParse(merge.Reference?.Value, out ExcelReference? merged) && merged!.Intersects(range))
                || WorksheetRoot.Descendants<CellFormula>().Any(formula =>
                    (formula.FormulaType?.Value == CellFormulaValues.Array || formula.FormulaType?.Value == CellFormulaValues.DataTable)
                    && ExcelReference.TryParse(formula.Reference?.Value, out ExcelReference? formulaRange) && formulaRange!.Intersects(range))
                || _worksheetPart.PivotTableParts.Any(part => ExcelReference.TryParse(part.PivotTableDefinition?.Location?.Reference?.Value, out ExcelReference? pivot) && pivot!.Intersects(range));
            if (conflict) throw new InvalidOperationException(message);
        }

        private void ApplyCellShift(ExcelReference affected, ExcelCellShiftDirection direction, bool inserting, CancellationToken cancellationToken) {
            affected.GetBounds(out int r1, out int c1, out int r2, out int c2);
            int rows = r2 - r1 + 1;
            int columns = c2 - c1 + 1;
            List<(Cell Cell, int Row, int Column)> cells = WorksheetRoot.Descendants<Cell>()
                .Select(cell => TryGetCellCoordinates(cell, out int row, out int column) ? (cell, row, column) : (cell, 0, 0))
                .Where(item => item.Item2 > 0).ToList();
            IEnumerable<(Cell Cell, int Row, int Column)> ordered = direction is ExcelCellShiftDirection.Right or ExcelCellShiftDirection.Down
                ? cells.OrderByDescending(item => direction == ExcelCellShiftDirection.Right ? item.Column : item.Row)
                : cells.OrderBy(item => direction == ExcelCellShiftDirection.Left ? item.Column : item.Row);
            foreach ((Cell cell, int row, int column) in ordered) {
                cancellationToken.ThrowIfCancellationRequested();
                bool inRowBand = row >= r1 && row <= r2;
                bool inColumnBand = column >= c1 && column <= c2;
                if (direction == ExcelCellShiftDirection.Right && inRowBand && column >= c1) cell.CellReference = A1.CellReference(row, column + columns);
                else if (direction == ExcelCellShiftDirection.Down && inColumnBand && row >= r1) cell.CellReference = A1.CellReference(row + rows, column);
                else if (direction == ExcelCellShiftDirection.Left && inRowBand) {
                    if (column >= c1 && column <= c2) cell.Remove();
                    else if (column > c2) cell.CellReference = A1.CellReference(row, column - columns);
                } else if (direction == ExcelCellShiftDirection.Up && inColumnBand) {
                    if (row >= r1 && row <= r2) cell.Remove();
                    else if (row > r2) cell.CellReference = A1.CellReference(row - rows, column);
                }
            }
            foreach (Row row in WorksheetRoot.Descendants<Row>().Where(row => !row.Elements<Cell>().Any()).ToList()) row.Remove();
            _excelDocument.RewriteCellShiftReferences(this, affected, direction, inserting);
            RewriteDrawingCellShift(affected, direction, inserting);
            _excelDocument.CleanupCalculationArtifacts(save: false, ExcelCalculationCleanupPolicy.RequestFullCalculationOnOpen);
            ResetMutationCaches();
        }

        private void ApplyRangeTransfer(ExcelReference source, int destinationRow, int destinationColumn, bool move, bool transpose, CancellationToken cancellationToken) {
            source.GetBounds(out int sr1, out int sc1, out int sr2, out int sc2);
            int sourceRows = sr2 - sr1 + 1;
            int sourceColumns = sc2 - sc1 + 1;
            int destinationRows = transpose ? sourceColumns : sourceRows;
            int destinationColumns = transpose ? sourceRows : sourceColumns;
            var snapshots = WorksheetRoot.Descendants<Cell>()
                .Where(cell => TryGetCellCoordinates(cell, out int row, out int column) && source.Contains(row, column))
                .Select(cell => {
                    TryGetCellCoordinates(cell, out int row, out int column);
                    return (Row: row, Column: column, Cell: (Cell)cell.CloneNode(true));
                }).ToList();
            var imageSnapshots = Images.Where(image => source.Contains(image.RowIndex, image.ColumnIndex))
                .Select(image => new { Image = image, Bytes = image.ToBytes(), image.ContentType, image.WidthPixels, image.HeightPixels, image.OffsetXPixels, image.OffsetYPixels })
                .ToList();

            if (move) _excelDocument.RewriteMovedRangeReferences(this, source, destinationRow, destinationColumn, transpose);
            RemoveCellsInRange(destinationRow, destinationColumn, destinationRow + destinationRows - 1, destinationColumn + destinationColumns - 1);
            if (move) RemoveCellsInRange(sr1, sc1, sr2, sc2);
            foreach (var snapshot in snapshots) {
                cancellationToken.ThrowIfCancellationRequested();
                int rowOffset = snapshot.Row - sr1;
                int columnOffset = snapshot.Column - sc1;
                int targetRow = destinationRow + (transpose ? columnOffset : rowOffset);
                int targetColumn = destinationColumn + (transpose ? rowOffset : columnOffset);
                Cell clone = snapshot.Cell;
                clone.CellReference = A1.CellReference(targetRow, targetColumn);
                if (clone.CellFormula != null) {
                    clone.CellFormula.Text = move
                        ? TranslateMovedFormula(clone.CellFormula.Text, source, destinationRow, destinationColumn, transpose)
                        : TranslateCopiedFormula(clone.CellFormula.Text, snapshot.Row, snapshot.Column, targetRow, targetColumn, transpose);
                    clone.CellFormula.CalculateCell = true;
                    clone.CellValue = null;
                }
                PutClonedCell(targetRow, targetColumn, clone);
            }
            foreach (var image in imageSnapshots) {
                int rowOffset = image.Image.RowIndex - sr1;
                int columnOffset = image.Image.ColumnIndex - sc1;
                int targetRow = destinationRow + (transpose ? columnOffset : rowOffset);
                int targetColumn = destinationColumn + (transpose ? rowOffset : columnOffset);
                if (move) image.Image.MoveTo(targetRow, targetColumn, image.OffsetXPixels, image.OffsetYPixels);
                else AddImage(targetRow, targetColumn, image.Bytes, image.ContentType,
                    transpose ? image.HeightPixels : image.WidthPixels,
                    transpose ? image.WidthPixels : image.HeightPixels,
                    image.OffsetXPixels,
                    image.OffsetYPixels);
            }
            _excelDocument.CleanupCalculationArtifacts(save: false, ExcelCalculationCleanupPolicy.RequestFullCalculationOnOpen);
            ResetMutationCaches();
        }

        private string TranslateCopiedFormula(string formula, int sourceRow, int sourceColumn, int targetRow, int targetColumn, bool transpose) {
            return ExcelFormulaSyntaxTree.Parse(formula).Rewrite(reference => {
                int MapRow(ExcelReferencePoint point) => point.RowAbsolute ? point.Row
                    : transpose ? targetRow + point.Column - sourceColumn : targetRow + point.Row - sourceRow;
                int MapColumn(ExcelReferencePoint point) => point.ColumnAbsolute ? point.Column
                    : transpose ? targetColumn + point.Row - sourceRow : targetColumn + point.Column - sourceColumn;
                ExcelReferenceKind kind = transpose
                    ? reference.Kind == ExcelReferenceKind.WholeRow ? ExcelReferenceKind.WholeColumn
                        : reference.Kind == ExcelReferenceKind.WholeColumn ? ExcelReferenceKind.WholeRow
                        : reference.Kind
                    : reference.Kind;
                return reference.WithCoordinates(
                    kind,
                    kind == ExcelReferenceKind.WholeColumn ? 0 : MapRow(reference.Start),
                    kind == ExcelReferenceKind.WholeRow ? 0 : MapColumn(reference.Start),
                    kind == ExcelReferenceKind.WholeColumn ? 0 : MapRow(reference.End),
                    kind == ExcelReferenceKind.WholeRow ? 0 : MapColumn(reference.End),
                    transpose ? reference.Start.ColumnAbsolute : null,
                    transpose ? reference.Start.RowAbsolute : null,
                    transpose ? reference.End.ColumnAbsolute : null,
                    transpose ? reference.End.RowAbsolute : null);
            });
        }

        private string TranslateMovedFormula(string formula, ExcelReference source, int destinationRow, int destinationColumn, bool transpose) {
            source.GetBounds(out int sr1, out int sc1, out int sr2, out int sc2);
            return ExcelFormulaSyntaxTree.Parse(formula).Rewrite(reference => {
                if (!string.IsNullOrWhiteSpace(reference.Qualifier)
                    && !IsCurrentSheetQualifier(reference.Qualifier!, Name)) return reference;
                reference.GetBounds(out int rr1, out int rc1, out int rr2, out int rc2);
                if (rr1 < sr1 || rr2 > sr2 || rc1 < sc1 || rc2 > sc2) return reference;
                int MapRow(int row, int column) => transpose ? destinationRow + column - sc1 : destinationRow + row - sr1;
                int MapColumn(int row, int column) => transpose ? destinationColumn + row - sr1 : destinationColumn + column - sc1;
                ExcelReferenceKind kind = transpose
                    ? reference.Kind == ExcelReferenceKind.WholeRow ? ExcelReferenceKind.WholeColumn
                        : reference.Kind == ExcelReferenceKind.WholeColumn ? ExcelReferenceKind.WholeRow
                        : reference.Kind
                    : reference.Kind;
                return reference.WithCoordinates(
                    kind,
                    kind == ExcelReferenceKind.WholeColumn ? 0 : MapRow(reference.Start.Row, reference.Start.Column),
                    kind == ExcelReferenceKind.WholeRow ? 0 : MapColumn(reference.Start.Row, reference.Start.Column),
                    kind == ExcelReferenceKind.WholeColumn ? 0 : MapRow(reference.End.Row, reference.End.Column),
                    kind == ExcelReferenceKind.WholeRow ? 0 : MapColumn(reference.End.Row, reference.End.Column),
                    transpose ? reference.Start.ColumnAbsolute : null,
                    transpose ? reference.Start.RowAbsolute : null,
                    transpose ? reference.End.ColumnAbsolute : null,
                    transpose ? reference.End.RowAbsolute : null);
            });
        }

        private void PutClonedCell(int row, int column, Cell clone) {
            Cell target = GetCell(row, column);
            target.InsertBeforeSelf(clone);
            target.Remove();
        }

        private void RemoveCellsInRange(int r1, int c1, int r2, int c2) {
            foreach (Cell cell in WorksheetRoot.Descendants<Cell>().Where(cell =>
                TryGetCellCoordinates(cell, out int row, out int column)
                && row >= r1 && row <= r2 && column >= c1 && column <= c2).ToList()) cell.Remove();
            foreach (Row row in WorksheetRoot.Descendants<Row>().Where(row => !row.Elements<Cell>().Any()).ToList()) row.Remove();
        }

        private static bool TryGetCellCoordinates(Cell cell, out int row, out int column) {
            row = column = 0;
            if (cell.CellReference?.Value is not string reference) return false;
            (row, column) = A1.ParseCellRef(reference);
            return row > 0 && column > 0;
        }

        private void RewriteDrawingCellShift(ExcelReference affected, ExcelCellShiftDirection direction, bool inserting) {
            Xdr.WorksheetDrawing? drawing = _worksheetPart.DrawingsPart?.WorksheetDrawing;
            if (drawing == null) return;
            affected.GetBounds(out int r1, out int c1, out int r2, out int c2);
            int rows = r2 - r1 + 1;
            int columns = c2 - c1 + 1;
            foreach (Xdr.MarkerType marker in drawing.Descendants<Xdr.MarkerType>()) {
                int row = int.TryParse(marker.RowId?.Text, out int parsedRow) ? parsedRow + 1 : 1;
                int column = int.TryParse(marker.ColumnId?.Text, out int parsedColumn) ? parsedColumn + 1 : 1;
                if (direction == ExcelCellShiftDirection.Right && row >= r1 && row <= r2 && column >= c1) column += columns;
                else if (direction == ExcelCellShiftDirection.Down && column >= c1 && column <= c2 && row >= r1) row += rows;
                else if (direction == ExcelCellShiftDirection.Left && row >= r1 && row <= r2 && column > c2) column -= columns;
                else if (direction == ExcelCellShiftDirection.Up && column >= c1 && column <= c2 && row > r2) row -= rows;
                marker.RowId!.Text = (row - 1).ToString(System.Globalization.CultureInfo.InvariantCulture);
                marker.ColumnId!.Text = (column - 1).ToString(System.Globalization.CultureInfo.InvariantCulture);
            }
            drawing.Save();
        }

        private static ExcelReference ParseLocalCellRange(string range) {
            ExcelReference parsed = ExcelReference.Parse(range);
            if (parsed.IsQualified || parsed.Kind is ExcelReferenceKind.WholeColumn or ExcelReferenceKind.WholeRow) {
                throw new ArgumentException("A local cell or rectangular range is required.", nameof(range));
            }
            return parsed;
        }
    }
}
