using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private bool _dynamicArrayWriteIndexInitialized;
        private readonly List<(int Top, int Left, int Bottom, int Right)> _dynamicArrayWriteRanges =
            new List<(int Top, int Left, int Bottom, int Right)>();

        private void InvalidateDynamicArrayWriteIndex() => _dynamicArrayWriteIndexInitialized = false;

        private void EnsureDynamicArrayWriteIndex() {
            if (_dynamicArrayWriteIndexInitialized) return;
            _dynamicArrayWriteRanges.Clear();
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
                    _dynamicArrayWriteRanges.Add((top, left, bottom, right));
                }
            }
            _dynamicArrayWriteIndexInitialized = true;
        }

        private void EnsureDynamicArrayCellWritable(int row, int column) {
            EnsureDynamicArrayWriteIndex();
            if (_dynamicArrayWriteRanges.Any(range => row >= range.Top && row <= range.Bottom
                && column >= range.Left && column <= range.Right))
                throw new InvalidOperationException($"Cell '{A1.CellReference(row, column)}' belongs to a dynamic array; clear its array formula first.");
        }

        private Cell GetWritableValueCell(int row, int column) {
            EnsureDynamicArrayCellWritable(row, column);
            return GetCell(row, column);
        }

        private void EnsureDynamicArrayRangeWritable(int top, int left, int bottom, int right) {
            EnsureDynamicArrayWriteIndex();
            if (_dynamicArrayWriteRanges.Any(range =>
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
                    }
            }
            FixedArraySheetIndex index = GetFixedArraySheetIndex();
            for (int row = plan.Owner.Top; row <= plan.Owner.Bottom; row++)
                for (int column = plan.Owner.Left; column <= plan.Owner.Right; column++) {
                    if (row == plan.Owner.Top && column == plan.Owner.Left) continue;
                    if (plan.Array != null && row <= plan.Bottom && column <= plan.Right) continue;
                    if (!index.Cells.TryGetValue(DynamicCellKey(row, column), out Cell? cell)) continue;
                    if (cell.CellFormula != null || cell.CellMetaIndex != null) continue;
                    cell.CellValue = null;
                    cell.InlineString = null;
                    cell.DataType = null;
                    cell.ValueMetaIndex = null;
                    if (cell.StyleIndex == null && !cell.HasChildren) cell.Remove();
                }
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
                                && IsDynamicSpillObstacle(owner, cell, row, column)) {
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

        private static bool IsDynamicSpillObstacle(FixedArrayOwner owner, Cell cell, int row, int column) {
            if (cell.CellFormula != null || cell.CellMetaIndex != null) return true;
            if (cell.ValueMetaIndex != null && cell.DataType?.Value != DocumentFormat.OpenXml.Spreadsheet.CellValues.Error) return true;
            if (row >= owner.Top && row <= owner.Bottom && column >= owner.Left && column <= owner.Right)
                return false;
            return cell.CellValue != null || cell.InlineString != null || cell.ValueMetaIndex != null;
        }

        private static void RetireOldDynamicChildren(ArrayCalculationContext context, DynamicArrayPlan plan) {
            FixedArrayOwner owner = plan.Owner;
            for (int row = owner.Top; row <= owner.Bottom; row++)
                for (int column = owner.Left; column <= owner.Right; column++) {
                    if (row == owner.Top && column == owner.Left) continue;
                    if (plan.Array != null && row <= plan.Bottom && column <= plan.Right) continue;
                    context.DynamicCells.Remove(DynamicCellKey(row, column));
                    context.DynamicRetiredCells.Add(DynamicCellKey(row, column));
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
