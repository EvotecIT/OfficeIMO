using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private static readonly System.Runtime.CompilerServices.ConditionalWeakTable<
            DocumentFormat.OpenXml.Packaging.WorksheetPart, DynamicSpillOwnership> DynamicSpillOwnerships =
            new System.Runtime.CompilerServices.ConditionalWeakTable<
                DocumentFormat.OpenXml.Packaging.WorksheetPart, DynamicSpillOwnership>();

        private sealed class DynamicSpillOwnership {
            internal readonly Dictionary<long, DynamicSpillCacheSnapshot> Cells =
                new Dictionary<long, DynamicSpillCacheSnapshot>();
            internal bool WriteIndexInitialized;
            internal readonly List<(int Top, int Left, int Bottom, int Right)> WriteRanges =
                new List<(int Top, int Left, int Bottom, int Right)>();
        }

        private DynamicSpillOwnership SpillOwnership => DynamicSpillOwnerships.GetOrCreateValue(_worksheetPart);
        private Dictionary<long, DynamicSpillCacheSnapshot> DynamicSpillCacheSnapshots =>
            SpillOwnership.Cells;

        private sealed class DynamicSpillCacheSnapshot {
            internal DocumentFormat.OpenXml.Spreadsheet.CellValues? Type;
            internal string? Value;
            internal uint? ValueMetaIndex;
        }

        private static DynamicSpillCacheSnapshot SnapshotDynamicSpillCell(Cell cell) => new DynamicSpillCacheSnapshot {
            Type = cell.DataType?.Value,
            Value = cell.CellValue?.Text,
            ValueMetaIndex = cell.ValueMetaIndex?.Value
        };

        private static bool MatchesDynamicSpillSnapshot(Cell cell, DynamicSpillCacheSnapshot snapshot) =>
            cell.DataType?.Value == snapshot.Type
            && string.Equals(cell.CellValue?.Text, snapshot.Value, StringComparison.Ordinal)
            && cell.ValueMetaIndex?.Value == snapshot.ValueMetaIndex;

        private void InvalidateDynamicArrayWriteIndex() => SpillOwnership.WriteIndexInitialized = false;

        private void EnsureDynamicArrayWriteIndex() {
            DynamicSpillOwnership ownership = SpillOwnership;
            if (ownership.WriteIndexInitialized) return;
            ownership.WriteRanges.Clear();
            var part = _excelDocument.WorkbookPartRoot.CellMetadataPart;
            if (part != null) {
                if (!part.IsRootElementLoaded) ValidateInCellImageMetadataPart(part, "Cell metadata");
                Metadata? metadata = part.Metadata;
                foreach (Cell cell in WorksheetRoot.Descendants<Cell>()) {
                    if (cell.CellFormula?.FormulaType?.Value != CellFormulaValues.Array
                        || cell.CellMetaIndex == null
                        || CreateFormulaArrayInfo(cell, metadata)?.IsDynamic != true
                        || !TryFixedArrayBounds(cell.CellFormula.Reference?.Value ?? "",
                            out int top, out int left, out int bottom, out int right)) continue;
                    ownership.WriteRanges.Add((top, left, bottom, right));
                }
            }
            ownership.WriteIndexInitialized = true;
        }

        private void EnsureDynamicArrayCellWritable(int row, int column) {
            EnsureDynamicArrayWriteIndex();
            if (SpillOwnership.WriteRanges.Any(range => row >= range.Top && row <= range.Bottom
                && column >= range.Left && column <= range.Right))
                throw new InvalidOperationException($"Cell '{A1.CellReference(row, column)}' belongs to a dynamic array; clear its array formula first.");
        }

        private Cell GetWritableValueCell(int row, int column) {
            EnsureDynamicArrayCellWritable(row, column);
            return GetCell(row, column);
        }

        private void EnsureDynamicArrayRangeWritable(int top, int left, int bottom, int right) {
            EnsureDynamicArrayWriteIndex();
            if (SpillOwnership.WriteRanges.Any(range =>
                RangesOverlapInclusive((top, left, bottom, right), range)))
                throw new InvalidOperationException("The range overlaps a dynamic array; clear its array formula first.");
        }

        private sealed class DynamicArrayPlan {
            internal FixedArrayOwner Owner = null!;
            internal FormulaArrayValue? Array;
            internal int Top, Left, Bottom, Right;
            internal int ErrorSubtype;
            internal int ErrorColumns, ErrorRows;
        }

        private bool TryGetDynamicArrayPlan(Cell cell, out DynamicArrayPlan plan) {
            plan = null!;
            return _formulaEvaluationCache != null
                && ArrayCalculationContexts.GetOrCreateValue(_formulaEvaluationCache).DynamicPlans.TryGetValue(
                    GetFormulaEvaluationCacheKey(cell.CellReference!.Value!), out plan!);
        }

        private void PlanDynamicArrayOwners() {
            var context = ArrayCalculationContexts.GetOrCreateValue(_formulaEvaluationCache!);
            if (context.DynamicOwnersPlanned) return;
            context.PlanningDynamicOwners = true;
            try {
                foreach (FixedArrayOwner owner in GetFixedArraySheetIndex().DynamicOwners)
                    TryEvaluateFormulaCellValue(owner.Cell, out _);
                context.DynamicOwnersPlanned = true;
            } finally {
                context.PlanningDynamicOwners = false;
            }
        }

        private void WriteDynamicArrayFormulaCache(DynamicArrayPlan plan) {
            Cell anchor = plan.Owner.Cell;
            string? previousReference = anchor.CellFormula?.Reference?.Value;
            if (plan.ErrorSubtype != 0) {
                anchor.CellFormula!.Reference = anchor.CellReference!.Value!;
                TryWriteRichFormulaError(anchor, "#SPILL!", plan);
            } else {
                FormulaArrayValue array = plan.Array!;
                anchor.CellFormula!.Reference = A1.CellReference(plan.Top, plan.Left) + ":"
                    + A1.CellReference(plan.Bottom, plan.Right);
                for (int row = plan.Top; row <= plan.Bottom; row++)
                    for (int column = plan.Left; column <= plan.Right; column++) {
                        Cell cell = row == plan.Top && column == plan.Left ? anchor : GetCell(row, column);
                        cell.InlineString = null;
                        cell.ValueMetaIndex = null;
                        SetFormulaCachedValue(cell, array.Values[(row - plan.Top) * array.Columns + column - plan.Left]);
                        if (cell != anchor)
                            DynamicSpillCacheSnapshots[DynamicCellKey(row, column)] = SnapshotDynamicSpillCell(cell);
                    }
            }
            FixedArraySheetIndex index = GetFixedArraySheetIndex();
            for (int row = plan.Owner.Top; row <= plan.Owner.Bottom; row++)
                for (int column = plan.Owner.Left; column <= plan.Owner.Right; column++) {
                    if (row == plan.Owner.Top && column == plan.Owner.Left) continue;
                    if (plan.Array != null && row <= plan.Bottom && column <= plan.Right) continue;
                    if (!index.Cells.TryGetValue(DynamicCellKey(row, column), out Cell? cell)) continue;
                    long key = DynamicCellKey(row, column);
                    if (!IsOwnedDynamicSpillCachedCell(cell, key)) continue;
                    cell.CellValue = null;
                    cell.InlineString = null;
                    cell.DataType = null;
                    cell.ValueMetaIndex = null;
                    DynamicSpillCacheSnapshots.Remove(key);
                    if (cell.StyleIndex == null && !cell.HasChildren) cell.Remove();
                }
            if (!string.Equals(previousReference, anchor.CellFormula?.Reference?.Value, StringComparison.Ordinal))
                InvalidateDynamicArrayWriteIndex();
        }

        private bool TryPlanDynamicArray(FixedArrayOwner owner, FormulaArrayValue array,
            out FormulaArgumentValue result) {
            var context = ArrayCalculationContexts.GetOrCreateValue(_formulaEvaluationCache!);
            var index = GetFixedArraySheetIndex();
            int top = owner.Top, left = owner.Left;
            long bottom = (long)top + array.Rows - 1, right = (long)left + array.Columns - 1;
            var plan = new DynamicArrayPlan {
                Owner = owner, Top = top, Left = left,
                Bottom = bottom <= A1.MaxRows ? (int)bottom : top,
                Right = right <= A1.MaxColumns ? (int)right : left,
                ErrorColumns = right <= A1.MaxColumns ? array.Columns - 1 : 0,
                ErrorRows = bottom <= A1.MaxRows ? array.Rows - 1 : 0
            };
            if (bottom > A1.MaxRows || right > A1.MaxColumns) {
                plan.ErrorSubtype = 3;
            } else {
                var target = (top, left, plan.Bottom, plan.Right);
                if (WorksheetRoot.Elements<MergeCells>().SelectMany(item => item.Elements<MergeCell>())
                    .Any(merge => TryFixedArrayBounds(merge.Reference?.Value ?? "",
                        out int mr1, out int mc1, out int mr2, out int mc2)
                        && RangesOverlapInclusive(target, (mr1, mc1, mr2, mc2)))) {
                    plan.ErrorSubtype = 6;
                } else if (_worksheetPart.TableDefinitionParts.Any(part =>
                    TryFixedArrayBounds(part.Table?.Reference?.Value ?? "",
                        out int tr1, out int tc1, out int tr2, out int tc2)
                    && RangesOverlapInclusive(target, (tr1, tc1, tr2, tc2)))) {
                    plan.ErrorSubtype = 1;
                } else if (index.Owners.Any(fixedOwner => RangesOverlapInclusive(target,
                    (fixedOwner.Top, fixedOwner.Left, fixedOwner.Bottom, fixedOwner.Right)))) {
                    plan.ErrorSubtype = 1;
                } else {
                    for (int row = top; row <= plan.Bottom && plan.ErrorSubtype == 0; row++)
                        for (int column = left; column <= plan.Right; column++) {
                            if (row == top && column == left) continue;
                            long key = DynamicCellKey(row, column);
                            if (context.DynamicCells.ContainsKey(key)
                                || index.Cells.TryGetValue(key, out Cell? cell)
                                && IsDynamicSpillObstacle(owner, cell, row, column,
                                    array.Values[(row - top) * array.Columns + column - left])) {
                                plan.ErrorSubtype = 1;
                                break;
                            }
                        }
                }
                if (plan.ErrorSubtype == 0) {
                    for (int row = owner.Top; row <= owner.Bottom && plan.ErrorSubtype == 0; row++)
                        for (int column = owner.Left; column <= owner.Right; column++) {
                            if (row == owner.Top && column == owner.Left
                                || row <= plan.Bottom && column <= plan.Right) continue;
                            long key = DynamicCellKey(row, column);
                            if (index.Cells.TryGetValue(key, out Cell? oldCell)
                                && !IsOwnedDynamicSpillCachedCell(oldCell, key)) {
                                plan.ErrorSubtype = 1;
                                break;
                            }
                        }
                }
            }
            context.DynamicPlans[GetFormulaEvaluationCacheKey(owner.Cell.CellReference!.Value!)] = plan;
            if (plan.ErrorSubtype != 0) {
                result = FormulaArgumentValue.Error("#SPILL!");
                RetireOldDynamicChildren(context, plan);
                return true;
            }
            plan.Array = array;
            context.Results[GetFormulaEvaluationCacheKey(owner.Cell.CellReference!.Value!)] = array;
            for (int row = top; row <= plan.Bottom; row++)
                for (int column = left; column <= plan.Right; column++)
                    if (row != top || column != left) {
                        context.DynamicCells[DynamicCellKey(row, column)] =
                            array.Values[(row - top) * array.Columns + column - left];
                        context.DynamicRetiredCells.Remove(DynamicCellKey(row, column));
                    }
            RetireOldDynamicChildren(context, plan);
            result = array.Values[0];
            return true;
        }

        private bool IsDynamicSpillObstacle(FixedArrayOwner owner, Cell cell, int row, int column,
            FormulaArgumentValue expected) {
            if (row >= owner.Top && row <= owner.Bottom && column >= owner.Left && column <= owner.Right)
                return !IsOwnedDynamicSpillCachedCell(cell, DynamicCellKey(row, column), expected);
            return cell.CellValue != null || cell.InlineString != null || cell.ValueMetaIndex != null;
        }

        private bool IsOwnedDynamicSpillCachedCell(Cell cell, long key, FormulaArgumentValue? expected = null) {
            if (cell.CellFormula != null || cell.CellMetaIndex != null || cell.InlineString != null) return false;
            if (DynamicSpillCacheSnapshots.TryGetValue(key, out DynamicSpillCacheSnapshot? snapshot))
                return MatchesDynamicSpillSnapshot(cell, snapshot);
            if (cell.CellValue == null && cell.ValueMetaIndex == null) return true;
            if (expected == null) return false;
            FormulaArgumentValue value = expected.Value;
            if (value.IsError)
                return cell.DataType?.Value == DocumentFormat.OpenXml.Spreadsheet.CellValues.Error
                    && ResolveRichValueError(cell, cell.CellValue?.Text) == value.ErrorCode;
            if (cell.ValueMetaIndex != null) return false;
            if (value.IsBoolean)
                return cell.DataType?.Value == DocumentFormat.OpenXml.Spreadsheet.CellValues.Boolean
                    && cell.CellValue?.Text == (value.Number == 0 ? "0" : "1");
            if (value.Number.HasValue)
                return (cell.DataType == null
                        || cell.DataType.Value == DocumentFormat.OpenXml.Spreadsheet.CellValues.Number)
                    && double.TryParse(cell.CellValue?.Text, System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out double numeric)
                    && numeric == value.Number.Value;
            return value.Text != null && string.Equals(GetCellText(cell), value.Text, StringComparison.Ordinal);
        }

        private void RetireOldDynamicChildren(ArrayCalculationContext context, DynamicArrayPlan plan) {
            FixedArrayOwner owner = plan.Owner;
            FixedArraySheetIndex index = GetFixedArraySheetIndex();
            for (int row = owner.Top; row <= owner.Bottom; row++)
                for (int column = owner.Left; column <= owner.Right; column++) {
                    if (row == owner.Top && column == owner.Left) continue;
                    if (plan.Array != null && row <= plan.Bottom && column <= plan.Right) continue;
                    long key = DynamicCellKey(row, column);
                    if (index.Cells.TryGetValue(key, out Cell? oldCell)
                        && !IsOwnedDynamicSpillCachedCell(oldCell, key)) continue;
                    context.DynamicCells.Remove(key);
                    context.DynamicRetiredCells.Add(key);
                }
        }

        private bool TryResolveDynamicArrayChild(int row, int column, out FormulaArgumentValue result) {
            result = default;
            var context = ArrayCalculationContexts.GetOrCreateValue(_formulaEvaluationCache!);
            long key = DynamicCellKey(row, column);
            if (context.DynamicCells.TryGetValue(key, out result)) return true;
            if (context.PlanningDynamicOwners && !context.DynamicOwnersPlanned) {
                FixedArraySheetIndex index = GetFixedArraySheetIndex();
                bool occupied = index.Cells.TryGetValue(key, out Cell? existing)
                    && (existing.CellValue != null || existing.InlineString != null);
                bool oldSpillChild = index.DynamicOwners.Any(owner => row >= owner.Top && row <= owner.Bottom
                    && column >= owner.Left && column <= owner.Right);
                if (!occupied || oldSpillChild) {
                    for (int i = index.DynamicOwners.Count - 1; i >= 0; i--) {
                        FixedArrayOwner owner = index.DynamicOwners[i];
                        if (row < owner.Top || column < owner.Left
                            || (long)(row - owner.Top + 1) * (column - owner.Left + 1) > MaxResolvedFormulaRangeCells
                            || _formulaEvaluationStack?.Contains(GetFormulaEvaluationCacheKey(owner.Cell.CellReference!.Value!)) == true)
                            continue;
                        TryEvaluateFormulaCellValue(owner.Cell, out _);
                        if (context.DynamicCells.TryGetValue(key, out result)) return true;
                    }
                }
            }
            if (context.DynamicCells.TryGetValue(key, out result)) return true;
            if (context.DynamicRetiredCells.Contains(key)) {
                result = new FormulaArgumentValue(0, null);
                return true;
            }
            return false;
        }
    }
}
