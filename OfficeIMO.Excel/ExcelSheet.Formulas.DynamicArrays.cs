namespace OfficeIMO.Excel {
    public partial class ExcelSheet {

        private sealed class DynamicArrayPlan {
            internal FixedArrayOwner Owner = null!;
            internal FormulaArrayValue? Array;
            internal int Top, Left, Bottom, Right;
            internal int ErrorSubtype;
            internal int ErrorColumns, ErrorRows;
        }

        private bool TryGetDynamicArrayPlan(ExcelFormulaCell cell, out DynamicArrayPlan plan) {
            plan = null!;
            return _formulaEvaluationCache != null
                && ArrayCalculationContexts.GetOrCreateValue(_formulaEvaluationCache).DynamicPlans.TryGetValue(
                    GetFormulaEvaluationCacheKey(cell.Reference!), out plan!);
        }

        private void PlanDynamicArrayOwners() {
            var context = GetFixedArraySheetIndex();
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

        private bool TryPlanDynamicArray(FixedArrayOwner owner, FormulaArrayValue array,
            out FormulaArgumentValue result) {
            var context = ArrayCalculationContexts.GetOrCreateValue(_formulaEvaluationCache!);
            var index = GetFixedArraySheetIndex();
            TryCaptureOriginalDynamicSpillCaches();
            int top = owner.Top, left = owner.Left;
            long bottom = (long)top + array.Rows - 1, right = (long)left + array.Columns - 1;
            var plan = new DynamicArrayPlan {
                Owner = owner, Top = top, Left = left,
                Bottom = bottom <= ExcelA1.MaxRows ? (int)bottom : top,
                Right = right <= ExcelA1.MaxColumns ? (int)right : left,
                ErrorColumns = right <= ExcelA1.MaxColumns ? array.Columns - 1 : 0,
                ErrorRows = bottom <= ExcelA1.MaxRows ? array.Rows - 1 : 0
            };
            if (bottom > ExcelA1.MaxRows || right > ExcelA1.MaxColumns) {
                plan.ErrorSubtype = 3;
            } else {
                var target = (top, left, plan.Bottom, plan.Right);
                if (EnumerateMergedRanges()
                    .Any(merge => TryFixedArrayBounds(merge,
                        out int mr1, out int mc1, out int mr2, out int mc2)
                        && RangesOverlapInclusive(target, (mr1, mc1, mr2, mc2)))) {
                    plan.ErrorSubtype = 6;
                } else if (EnumerateTableRanges().Any(tableRange =>
                    TryFixedArrayBounds(tableRange,
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
                            if (index.DynamicCells.ContainsKey(key)
                                || index.Cells.TryGetValue(key, out ExcelFormulaCell cell)
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
                            if (index.Cells.TryGetValue(key, out ExcelFormulaCell oldCell)
                                && !IsOwnedDynamicSpillCachedCell(oldCell, key)) {
                                plan.ErrorSubtype = 1;
                                break;
                            }
                        }
                }
            }
            context.DynamicPlans[GetFormulaEvaluationCacheKey(owner.Cell.Reference!)] = plan;
            if (plan.ErrorSubtype != 0) {
                result = FormulaArgumentValue.Error("#SPILL!");
                RetireOldDynamicChildren(plan);
                return true;
            }
            plan.Array = array;
            context.Results[GetFormulaEvaluationCacheKey(owner.Cell.Reference!)] = array;
            for (int row = top; row <= plan.Bottom; row++)
                for (int column = left; column <= plan.Right; column++)
                    if (row != top || column != left) {
                        index.DynamicCells[DynamicCellKey(row, column)] =
                            array.Values[(row - top) * array.Columns + column - left];
                        index.DynamicRetiredCells.Remove(DynamicCellKey(row, column));
                    }
            RetireOldDynamicChildren(plan);
            result = array.Values[0];
            return true;
        }

        private void RetireOldDynamicChildren(DynamicArrayPlan plan) {
            FixedArrayOwner owner = plan.Owner;
            FixedArraySheetIndex index = GetFixedArraySheetIndex();
            for (int row = owner.Top; row <= owner.Bottom; row++)
                for (int column = owner.Left; column <= owner.Right; column++) {
                    if (row == owner.Top && column == owner.Left) continue;
                    if (plan.Array != null && row <= plan.Bottom && column <= plan.Right) continue;
                    long key = DynamicCellKey(row, column);
                    if (index.Cells.TryGetValue(key, out ExcelFormulaCell oldCell)
                        && !IsOwnedDynamicSpillCachedCell(oldCell, key)) continue;
                    index.DynamicCells.Remove(key);
                    index.DynamicRetiredCells.Add(key);
                }
        }

        private bool TryResolveDynamicArrayChild(int row, int column, out FormulaArgumentValue result) {
            result = default;
            var context = GetFixedArraySheetIndex();
            if (!context.PlanningDynamicOwners && !context.DynamicOwnersPlanned) PlanDynamicArrayOwners();
            long key = DynamicCellKey(row, column);
            if (context.DynamicCells.TryGetValue(key, out result)) return true;
            if (context.PlanningDynamicOwners && !context.DynamicOwnersPlanned) {
                bool occupied = context.Cells.TryGetValue(key, out ExcelFormulaCell existing) && existing.HasContent;
                bool oldSpillChild = false;
                foreach (FixedArrayOwner owner in context.DynamicOwners) {
                    if (row >= owner.Top && row <= owner.Bottom && column >= owner.Left && column <= owner.Right) { oldSpillChild = true; break; }
                }
                if (!occupied || oldSpillChild) {
                    for (int i = context.DynamicOwners.Count - 1; i >= 0; i--) {
                        FixedArrayOwner owner = context.DynamicOwners[i];
                        if (row < owner.Top || column < owner.Left
                            || (long)(row - owner.Top + 1) * (column - owner.Left + 1) > MaxResolvedFormulaRangeCells
                            || _formulaEvaluationStack?.Contains(GetFormulaEvaluationCacheKey(owner.Cell.Reference!)) == true)
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
