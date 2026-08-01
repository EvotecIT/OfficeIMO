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
            ExcelReference shiftedBand = GetCellShiftBand(affected, direction);
            return Locking.ExecuteRead(_excelDocument.EnsureLock(), () => {
                EnsureMutationPlanCanInspectWithoutMaterializing();
                MutationPlanScanBudget budget = CreateMutationPlanScanBudget(effective);
                ValidatePackageMutationReferenceSafety(
                    "Cell shifts",
                    budget.Consume,
                    shiftedBand,
                    direction,
                    reference => ExcelDocument.TransformCellShiftReference(
                        reference,
                        affected,
                        direction,
                        inserting));
                PreflightCellShift(affected, direction, inserting, budget);
                int count = InspectMutationPlanElements(WorksheetRoot.Descendants<Cell>(), budget).Count(cell => {
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
                        ExcelStructuralMutationPlan current = PlanCellShift(affected.ToString(), direction, inserting, effective);
                        ApplyCellShift(affected, direction, inserting, cancellationToken);
                        return current.AffectedCells;
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
                MutationPlanScanBudget budget = CreateMutationPlanScanBudget(effective);
                if (move) {
                    ValidatePackageMutationReferenceSafety(
                        "Range moves",
                        budget.Consume,
                        source);
                } else {
                    ValidateA1MutationReferenceMode("Range transfers");
                }
                PreflightRangeTransfer(source, destination.Start.Row, destination.Start.Column, destinationRows, destinationColumns, move, transpose, budget);
                int existing = InspectMutationPlanElements(WorksheetRoot.Descendants<Cell>(), budget).Count(cell =>
                    TryGetCellCoordinates(cell, out int row, out int column)
                    && source.Contains(row, column));
                int images = InspectMutationPlanElements(Images, budget)
                    .Count(image => !image.HasAbsoluteAnchor && source.Contains(image.RowIndex, image.ColumnIndex));
                var impacts = new List<ExcelMutationImpact> {
                    new ExcelMutationImpact("cells", existing, "Existing cells, formulas, and styles will be transferred.")
                };
                if (images > 0) impacts.Add(new ExcelMutationImpact("images", images, "Cell-anchored images will follow the transfer."));
                ExcelStructuralMutationKind kind = transpose ? ExcelStructuralMutationKind.Transpose : move ? ExcelStructuralMutationKind.Move : ExcelStructuralMutationKind.Copy;
                return new ExcelStructuralMutationPlan(this, kind, source.ToString(), destination.ToString(), (int)area, impacts, effective,
                    cancellationToken => {
                        cancellationToken.ThrowIfCancellationRequested();
                        ExcelStructuralMutationPlan current = PlanRangeTransfer(source.ToString(), destination.ToString(), move, transpose, effective);
                        ApplyRangeTransfer(
                            source,
                            destination.Start.Row,
                            destination.Start.Column,
                            move,
                            transpose,
                            effective.MaximumSnapshotCharacters,
                            cancellationToken);
                        return current.AffectedCells;
                    });
            });
        }

        private void PreflightCellShift(
            ExcelReference affected,
            ExcelCellShiftDirection direction,
            bool inserting,
            MutationPlanScanBudget? budget = null) {
            ValidateWorkbookSharedFormulasForStructuralEdit();
            ValidateStructuralVmlControlSafety(budget);
            affected.GetBounds(out int r1, out int c1, out int r2, out int c2);
            EnsureNoIntersectingOwnedStructures(
                GetCellShiftBand(affected, direction),
                "Cell shifts cannot split tables, merged cells, array formulas, data tables, or PivotTable output ranges in the shifted band.",
                budget);
            ValidateCellShiftConnectionParameters(affected, direction, inserting, budget);
            if (!inserting) return;
            if (direction == ExcelCellShiftDirection.Right) {
                int maxCell = InspectMutationPlanElements(WorksheetRoot.Descendants<Cell>(), budget).Where(cell => TryGetCellCoordinates(cell, out int row, out _) && row >= r1 && row <= r2)
                    .Select(cell => TryGetCellCoordinates(cell, out _, out int column) ? column : 0).DefaultIfEmpty().Max();
                int maxDrawing = InspectMutationPlanElements(
                        _worksheetPart.DrawingsPart?.WorksheetDrawing?.Descendants<Xdr.MarkerType>() ?? Enumerable.Empty<Xdr.MarkerType>(),
                        budget)
                    .Where(marker => ExcelDocument.DrawingMarkerMovesWithPlacement(marker, candidate =>
                        TryGetDrawingMarkerCoordinates(candidate, out int row, out int column)
                        && row >= r1 && row <= r2 && column >= c1))
                    .Select(marker => TryGetDrawingMarkerCoordinates(marker, out _, out int column) ? column : 0)
                    .DefaultIfEmpty().Max();
                int max = Math.Max(maxCell, maxDrawing);
                if ((long)max + c2 - c1 + 1L > A1.MaxColumns) throw new InvalidOperationException("Cell insertion would exceed the worksheet column limit.");
            } else {
                int maxCell = InspectMutationPlanElements(WorksheetRoot.Descendants<Cell>(), budget).Where(cell => TryGetCellCoordinates(cell, out _, out int column) && column >= c1 && column <= c2)
                    .Select(cell => TryGetCellCoordinates(cell, out int row, out _) ? row : 0).DefaultIfEmpty().Max();
                int maxDrawing = InspectMutationPlanElements(
                        _worksheetPart.DrawingsPart?.WorksheetDrawing?.Descendants<Xdr.MarkerType>() ?? Enumerable.Empty<Xdr.MarkerType>(),
                        budget)
                    .Where(marker => ExcelDocument.DrawingMarkerMovesWithPlacement(marker, candidate =>
                        TryGetDrawingMarkerCoordinates(candidate, out int row, out int column)
                        && column >= c1 && column <= c2 && row >= r1))
                    .Select(marker => TryGetDrawingMarkerCoordinates(marker, out int row, out _) ? row : 0)
                    .DefaultIfEmpty().Max();
                int max = Math.Max(maxCell, maxDrawing);
                if ((long)max + r2 - r1 + 1L > A1.MaxRows) throw new InvalidOperationException("Cell insertion would exceed the worksheet row limit.");
            }
        }

        private void PreflightRangeTransfer(
            ExcelReference source,
            int destinationRow,
            int destinationColumn,
            int rows,
            int columns,
            bool move,
            bool transpose,
            MutationPlanScanBudget? budget = null) {
            ValidateWorkbookSharedFormulasForStructuralEdit();
            ExcelReference destination = ExcelReference.Parse(A1.CellReference(destinationRow, destinationColumn) + ":" + A1.CellReference(destinationRow + rows - 1, destinationColumn + columns - 1));
            if (move || transpose) {
                EnsureNoIntersectingOwnedStructures(source, "Move and transpose cannot split owned table, merge, array, data-table, or PivotTable structures.", budget);
            }
            EnsureNoIntersectingOwnedStructures(destination, "Range-transfer destination cannot overwrite owned table, merge, array, data-table, or PivotTable structures.", budget);
            if (move) ValidateRangeMoveHyperlinks(source, destination, budget);
            foreach (Cell cell in InspectMutationPlanElements(WorksheetRoot.Descendants<Cell>(), budget)
                .Where(cell => TryGetCellCoordinates(cell, out int row, out int column) && source.Contains(row, column))) {
                CellFormulaValues? type = cell.CellFormula?.FormulaType?.Value;
                if (type == CellFormulaValues.Array || type == CellFormulaValues.DataTable) {
                    throw new InvalidOperationException("Range transfer cannot split array or data-table formulas.");
                }
            }
        }

        private void EnsureNoIntersectingOwnedStructures(
            ExcelReference range,
            string message,
            MutationPlanScanBudget? budget = null,
            Table? excludedTable = null) {
            bool conflict = InspectMutationPlanElements(_worksheetPart.TableDefinitionParts, budget)
                    .Any(part => !ReferenceEquals(part.Table, excludedTable)
                        && ExcelReference.TryParse(part.Table?.Reference?.Value, out ExcelReference? table)
                        && table!.Intersects(range))
                || InspectMutationPlanElements(WorksheetRoot.Descendants<MergeCell>(), budget)
                    .Any(merge => ExcelReference.TryParse(merge.Reference?.Value, out ExcelReference? merged) && merged!.Intersects(range))
                || InspectMutationPlanElements(WorksheetRoot.Descendants<CellFormula>(), budget).Any(formula =>
                    (formula.FormulaType?.Value == CellFormulaValues.Array || formula.FormulaType?.Value == CellFormulaValues.DataTable)
                    && ExcelReference.TryParse(formula.Reference?.Value, out ExcelReference? formulaRange) && formulaRange!.Intersects(range))
                || InspectMutationPlanElements(_worksheetPart.PivotTableParts, budget)
                    .Any(part => ExcelReference.TryParse(part.PivotTableDefinition?.Location?.Reference?.Value, out ExcelReference? pivot) && pivot!.Intersects(range));
            if (conflict) throw new InvalidOperationException(message);
        }

        private static ExcelReference GetCellShiftBand(
            ExcelReference affected,
            ExcelCellShiftDirection direction) {
            affected.GetBounds(out int r1, out int c1, out int r2, out int c2);
            string reference = direction == ExcelCellShiftDirection.Right || direction == ExcelCellShiftDirection.Left
                ? A1.CellReference(r1, c1) + ":" + A1.CellReference(r2, A1.MaxColumns)
                : A1.CellReference(r1, c1) + ":" + A1.CellReference(A1.MaxRows, c2);
            return ExcelReference.Parse(reference);
        }

        private void ApplyCellShift(ExcelReference affected, ExcelCellShiftDirection direction, bool inserting, CancellationToken cancellationToken) {
            affected.GetBounds(out int r1, out int c1, out int r2, out int c2);
            int rows = r2 - r1 + 1;
            int columns = c2 - c1 + 1;
            MaterializeWorkbookSharedFormulasForStructuralEdit();
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
                else if (direction == ExcelCellShiftDirection.Down && inColumnBand && row >= r1) MoveCellTo(cell, row + rows, column);
                else if (direction == ExcelCellShiftDirection.Left && inRowBand) {
                    if (column >= c1 && column <= c2) cell.Remove();
                    else if (column > c2) cell.CellReference = A1.CellReference(row, column - columns);
                } else if (direction == ExcelCellShiftDirection.Up && inColumnBand) {
                    if (row >= r1 && row <= r2) cell.Remove();
                    else if (row > r2) MoveCellTo(cell, row - rows, column);
                }
            }
            foreach (Row row in WorksheetRoot.Descendants<Row>().Where(row => !row.Elements<Cell>().Any()).ToList()) row.Remove();
            _excelDocument.RewriteCellShiftReferences(this, affected, direction, inserting);
            RewriteDrawingCellShift(affected, direction, inserting);
            _excelDocument.CleanupCalculationArtifacts(save: false, ExcelCalculationCleanupPolicy.RequestFullCalculationOnOpen);
            ResetMutationCaches();
        }

        private void MoveCellTo(Cell cell, int row, int column) {
            cell.Remove();
            cell.CellReference = A1.CellReference(row, column);
            Cell target = GetCell(row, column);
            target.InsertBeforeSelf(cell);
            target.Remove();
        }

        private void ApplyRangeTransfer(
            ExcelReference source,
            int destinationRow,
            int destinationColumn,
            bool move,
            bool transpose,
            long maximumImageSnapshotBytes,
            CancellationToken cancellationToken) {
            source.GetBounds(out int sr1, out int sc1, out int sr2, out int sc2);
            int sourceRows = sr2 - sr1 + 1;
            int sourceColumns = sc2 - sc1 + 1;
            int destinationRows = transpose ? sourceColumns : sourceRows;
            int destinationColumns = transpose ? sourceRows : sourceColumns;
            MaterializeWorkbookSharedFormulasForStructuralEdit();
            var snapshots = WorksheetRoot.Descendants<Cell>()
                .Where(cell => TryGetCellCoordinates(cell, out int row, out int column) && source.Contains(row, column))
                .Select(cell => {
                    TryGetCellCoordinates(cell, out int row, out int column);
                    return (Row: row, Column: column, Cell: (Cell)cell.CloneNode(true));
                }).ToList();
            List<RangeTransferImageSnapshot> imageSnapshots = CaptureRangeTransferImageSnapshots(
                Images.Where(image => !image.HasAbsoluteAnchor && source.Contains(image.RowIndex, image.ColumnIndex)),
                move,
                maximumImageSnapshotBytes);

            if (move) {
                RemoveRangeMoveDestinationComments(
                    source,
                    destinationRow,
                    destinationColumn,
                    destinationRow + destinationRows - 1,
                    destinationColumn + destinationColumns - 1);
                RemoveRangeMoveDestinationHyperlinks(
                    source,
                    destinationRow,
                    destinationColumn,
                    destinationRow + destinationRows - 1,
                    destinationColumn + destinationColumns - 1);
                _excelDocument.RewriteMovedRangeReferences(this, source, destinationRow, destinationColumn, transpose);
            }
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
                cancellationToken.ThrowIfCancellationRequested();
                int rowOffset = image.RowIndex - sr1;
                int columnOffset = image.ColumnIndex - sc1;
                int targetRow = destinationRow + (transpose ? columnOffset : rowOffset);
                int targetColumn = destinationColumn + (transpose ? rowOffset : columnOffset);
                if (move) {
                    image.Image.MoveTo(targetRow, targetColumn, image.OffsetXPixels, image.OffsetYPixels);
                    continue;
                }

                ExcelImage copy;
                if (image.HasTwoCellAnchor && image.ToRowIndex.HasValue && image.ToColumnIndex.HasValue) {
                    int toRow = transpose
                        ? targetRow + image.ToColumnIndex.Value - image.ColumnIndex
                        : targetRow + image.ToRowIndex.Value - image.RowIndex;
                    int toColumn = transpose
                        ? targetColumn + image.ToRowIndex.Value - image.RowIndex
                        : targetColumn + image.ToColumnIndex.Value - image.ColumnIndex;
                    if (toRow > A1.MaxRows || toColumn > A1.MaxColumns) {
                        throw new InvalidOperationException("Copied image anchor exceeds worksheet limits.");
                    }
                    copy = AddImageToRange(
                        A1.CellReference(targetRow, targetColumn) + ":" + A1.CellReference(targetRow, targetColumn),
                        image.Bytes!,
                        image.ContentType,
                        transpose ? image.OffsetYPixels : image.OffsetXPixels,
                        transpose ? image.OffsetXPixels : image.OffsetYPixels,
                        name: image.Name,
                        altText: image.Description,
                        title: image.Title,
                        lockAspectRatio: image.IsAspectRatioLocked,
                        placement: image.Placement,
                        rotationDegrees: image.RotationDegrees);
                    copy.SetTwoCellEndingMarker(
                        toRow,
                        toColumn,
                        transpose ? image.ToOffsetYPixels : image.ToOffsetXPixels,
                        transpose ? image.ToOffsetXPixels : image.ToOffsetYPixels);
                    copy.SetSize(
                        transpose ? image.HeightPixels : image.WidthPixels,
                        transpose ? image.WidthPixels : image.HeightPixels);
                } else {
                    copy = AddImage(
                        targetRow,
                        targetColumn,
                        image.Bytes!,
                        image.ContentType,
                        transpose ? image.HeightPixels : image.WidthPixels,
                        transpose ? image.WidthPixels : image.HeightPixels,
                        transpose ? image.OffsetYPixels : image.OffsetXPixels,
                        transpose ? image.OffsetXPixels : image.OffsetYPixels,
                        image.Name,
                        image.Description,
                        image.IsAspectRatioLocked);
                    copy.Title = image.Title;
                    copy.SetRotation(image.RotationDegrees);
                }
                copy.SetCropRatio(image.CropLeftRatio, image.CropTopRatio, image.CropRightRatio, image.CropBottomRatio);
                copy.SetFlip(image.FlipHorizontal, image.FlipVertical);
            }
            _excelDocument.CleanupCalculationArtifacts(save: false, ExcelCalculationCleanupPolicy.RequestFullCalculationOnOpen);
            ResetMutationCaches();
        }

        private string TranslateCopiedFormula(string formula, int sourceRow, int sourceColumn, int targetRow, int targetColumn, bool transpose) {
            return ExcelFormulaSyntaxTree.Parse(formula).Rewrite(reference => {
                int MapRow(ExcelReferencePoint point) => transpose
                    ? point.ColumnAbsolute ? point.Column : targetRow + point.Column - sourceColumn
                    : point.RowAbsolute ? point.Row : targetRow + point.Row - sourceRow;
                int MapColumn(ExcelReferencePoint point) => transpose
                    ? point.RowAbsolute ? point.Row : targetColumn + point.Row - sourceRow
                    : point.ColumnAbsolute ? point.Column : targetColumn + point.Column - sourceColumn;
                ExcelReferenceKind kind = transpose
                    ? reference.Kind == ExcelReferenceKind.WholeRow ? ExcelReferenceKind.WholeColumn
                        : reference.Kind == ExcelReferenceKind.WholeColumn ? ExcelReferenceKind.WholeRow
                        : reference.Kind
                    : reference.Kind;
                int startRow = kind == ExcelReferenceKind.WholeColumn ? 0 : MapRow(reference.Start);
                int startColumn = kind == ExcelReferenceKind.WholeRow ? 0 : MapColumn(reference.Start);
                int endRow = kind == ExcelReferenceKind.WholeColumn ? 0 : MapRow(reference.End);
                int endColumn = kind == ExcelReferenceKind.WholeRow ? 0 : MapColumn(reference.End);
                if (!CoordinatesFitWorksheet(kind, startRow, startColumn, endRow, endColumn)) return null;
                return reference.WithCoordinates(
                    kind,
                    startRow,
                    startColumn,
                    endRow,
                    endColumn,
                    transpose ? reference.Start.ColumnAbsolute : null,
                    transpose ? reference.Start.RowAbsolute : null,
                    transpose ? reference.End.ColumnAbsolute : null,
                    transpose ? reference.End.RowAbsolute : null);
            });
        }

        private static bool CoordinatesFitWorksheet(
            ExcelReferenceKind kind,
            int startRow,
            int startColumn,
            int endRow,
            int endColumn) {
            bool rowsFit = kind == ExcelReferenceKind.WholeColumn
                || (startRow >= 1 && startRow <= A1.MaxRows && endRow >= 1 && endRow <= A1.MaxRows);
            bool columnsFit = kind == ExcelReferenceKind.WholeRow
                || (startColumn >= 1 && startColumn <= A1.MaxColumns && endColumn >= 1 && endColumn <= A1.MaxColumns);
            return rowsFit && columnsFit;
        }

        private static List<RangeTransferImageSnapshot> CaptureRangeTransferImageSnapshots(
            IEnumerable<ExcelImage> images,
            bool move,
            long maximumBytes) {
            var snapshots = new List<RangeTransferImageSnapshot>();
            long remainingBytes = maximumBytes;
            foreach (ExcelImage image in images) {
                byte[]? bytes = null;
                if (!move) {
                    if (!image.TryReadBytes(remainingBytes, out byte[] capturedBytes)) {
                        throw new InvalidOperationException(
                            $"Range-transfer image snapshots exceed MaximumSnapshotCharacters ({maximumBytes}).");
                    }
                    bytes = capturedBytes;
                    remainingBytes = checked(remainingBytes - capturedBytes.LongLength);
                }
                snapshots.Add(new RangeTransferImageSnapshot(image, bytes));
            }
            return snapshots;
        }

        private sealed class RangeTransferImageSnapshot {
            internal RangeTransferImageSnapshot(ExcelImage image, byte[]? bytes) {
                Image = image;
                Bytes = bytes;
                ContentType = image.ContentType;
                Name = image.Name;
                Title = image.Title;
                Description = image.Description;
                IsAspectRatioLocked = image.IsAspectRatioLocked;
                RowIndex = image.RowIndex;
                ColumnIndex = image.ColumnIndex;
                WidthPixels = image.WidthPixels;
                HeightPixels = image.HeightPixels;
                OffsetXPixels = image.OffsetXPixels;
                OffsetYPixels = image.OffsetYPixels;
                HasTwoCellAnchor = image.HasTwoCellAnchor;
                ToRowIndex = image.ToRowIndex;
                ToColumnIndex = image.ToColumnIndex;
                ToOffsetXPixels = image.ToOffsetXPixels;
                ToOffsetYPixels = image.ToOffsetYPixels;
                Placement = image.Placement;
                CropLeftRatio = image.CropLeftRatio;
                CropTopRatio = image.CropTopRatio;
                CropRightRatio = image.CropRightRatio;
                CropBottomRatio = image.CropBottomRatio;
                RotationDegrees = image.RotationDegrees;
                FlipHorizontal = image.FlipHorizontal;
                FlipVertical = image.FlipVertical;
            }

            internal ExcelImage Image { get; }
            internal byte[]? Bytes { get; }
            internal string ContentType { get; }
            internal string Name { get; }
            internal string Title { get; }
            internal string Description { get; }
            internal bool IsAspectRatioLocked { get; }
            internal int RowIndex { get; }
            internal int ColumnIndex { get; }
            internal int WidthPixels { get; }
            internal int HeightPixels { get; }
            internal int OffsetXPixels { get; }
            internal int OffsetYPixels { get; }
            internal bool HasTwoCellAnchor { get; }
            internal int? ToRowIndex { get; }
            internal int? ToColumnIndex { get; }
            internal int ToOffsetXPixels { get; }
            internal int ToOffsetYPixels { get; }
            internal ExcelImagePlacement Placement { get; }
            internal double CropLeftRatio { get; }
            internal double CropTopRatio { get; }
            internal double CropRightRatio { get; }
            internal double CropBottomRatio { get; }
            internal double RotationDegrees { get; }
            internal bool FlipHorizontal { get; }
            internal bool FlipVertical { get; }
        }

        private string TranslateMovedFormula(string formula, ExcelReference source, int destinationRow, int destinationColumn, bool transpose) {
            return ExcelFormulaSyntaxTree.Parse(formula).Rewrite(reference => {
                if (!string.IsNullOrWhiteSpace(reference.Qualifier)
                    && !IsCurrentSheetQualifier(reference.Qualifier!, Name)) return reference;
                return ExcelDocument.TransformMovedRangeReference(reference, source, destinationRow, destinationColumn, transpose);
            });
        }

        internal void RemapMovedConnectionParameters(
            ExcelReference source,
            int destinationRow,
            int destinationColumn,
            bool transpose) {
            Connections? connections = WorkbookPartRoot.ConnectionsPart?.Connections;
            if (connections == null) return;

            HashSet<uint> connectionIds = GetWorksheetQueryConnectionIds(_worksheetPart);
            bool changed = false;
            foreach (Connection connection in connections.Elements<Connection>()) {
                foreach (Parameter parameter in connection.Descendants<Parameter>()) {
                    if (parameter.Cell?.Value is not string value
                        || !ExcelReference.TryParse(value, out ExcelReference? reference)
                        || !ConnectionParameterTargetsCurrentSheet(connection, reference!, connectionIds)) continue;
                    ExcelReference mapped = ExcelDocument.TransformMovedRangeReference(
                        reference!, source, destinationRow, destinationColumn, transpose);
                    string rewritten = mapped.ToString();
                    if (string.Equals(value, rewritten, StringComparison.OrdinalIgnoreCase)) continue;
                    parameter.Cell = rewritten;
                    changed = true;
                }
            }
            if (changed) connections.Save();
        }

        private void ValidateCellShiftConnectionParameters(
            ExcelReference affected,
            ExcelCellShiftDirection direction,
            bool inserting,
            MutationPlanScanBudget? budget) {
            Connections? connections = WorkbookPartRoot.ConnectionsPart?.Connections;
            if (connections == null) return;

            HashSet<uint> connectionIds = GetWorksheetQueryConnectionIds(_worksheetPart, budget);
            foreach (Connection connection in InspectMutationPlanElements(connections.Elements<Connection>(), budget)) {
                foreach (Parameter parameter in InspectMutationPlanElements(connection.Descendants<Parameter>(), budget)) {
                    if (parameter.Cell?.Value is not string value
                        || !ExcelReference.TryParse(value, out ExcelReference? reference)
                        || !ConnectionParameterTargetsCurrentSheet(connection, reference!, connectionIds)) continue;
                    ExcelReference? mapped;
                    try {
                        mapped = ExcelDocument.TransformCellShiftReference(reference!, affected, direction, inserting);
                    } catch (Exception exception) when (exception is OverflowException || exception is ArgumentOutOfRangeException) {
                        throw new InvalidOperationException(
                            $"Cell insertion would move cell-backed connection parameter '{value}' beyond worksheet limits.",
                            exception);
                    }
                    if (mapped == null) {
                        throw new InvalidOperationException(
                            $"Cannot delete cell-backed connection parameter reference '{value}'. Update or remove the parameter first.");
                    }
                }
            }
        }

        internal void RemapCellShiftConnectionParameters(
            ExcelReference affected,
            ExcelCellShiftDirection direction,
            bool inserting) {
            Connections? connections = WorkbookPartRoot.ConnectionsPart?.Connections;
            if (connections == null) return;

            HashSet<uint> connectionIds = GetWorksheetQueryConnectionIds(_worksheetPart);
            bool changed = false;
            foreach (Connection connection in connections.Elements<Connection>()) {
                foreach (Parameter parameter in connection.Descendants<Parameter>()) {
                    if (parameter.Cell?.Value is not string value
                        || !ExcelReference.TryParse(value, out ExcelReference? reference)
                        || !ConnectionParameterTargetsCurrentSheet(connection, reference!, connectionIds)) continue;
                    ExcelReference? mapped = ExcelDocument.TransformCellShiftReference(
                        reference!, affected, direction, inserting);
                    if (mapped == null) {
                        throw new InvalidOperationException(
                            $"Cannot delete cell-backed connection parameter reference '{value}'. Update or remove the parameter first.");
                    }
                    string rewritten = mapped.ToString();
                    if (string.Equals(value, rewritten, StringComparison.OrdinalIgnoreCase)) continue;
                    parameter.Cell = rewritten;
                    changed = true;
                }
            }
            if (changed) connections.Save();
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

        private static bool TryGetDrawingMarkerCoordinates(Xdr.MarkerType marker, out int row, out int column) {
            row = column = 0;
            if (!int.TryParse(marker.RowId?.Text, out int zeroBasedRow)
                || !int.TryParse(marker.ColumnId?.Text, out int zeroBasedColumn)) return false;
            row = zeroBasedRow + 1;
            column = zeroBasedColumn + 1;
            return row > 0 && column > 0;
        }

        private void RewriteDrawingCellShift(ExcelReference affected, ExcelCellShiftDirection direction, bool inserting) {
            affected.GetBounds(out int r1, out int c1, out int r2, out int c2);
            int rows = r2 - r1 + 1;
            int columns = c2 - c1 + 1;
            ExcelDocument.RewriteDrawingAnchors(
                _worksheetPart.DrawingsPart?.WorksheetDrawing,
                (row, column) => {
                    if (direction == ExcelCellShiftDirection.Right && row >= r1 && row <= r2 && column >= c1) column += columns;
                    else if (direction == ExcelCellShiftDirection.Down && column >= c1 && column <= c2 && row >= r1) row += rows;
                    else if (direction == ExcelCellShiftDirection.Left && row >= r1 && row <= r2) {
                        if (column > c2) column -= columns;
                        else if (!inserting && column >= c1) column = c1;
                    } else if (direction == ExcelCellShiftDirection.Up && column >= c1 && column <= c2) {
                        if (row > r2) row -= rows;
                        else if (!inserting && row >= r1) row = r1;
                    }
                    return (row, column);
                });
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
