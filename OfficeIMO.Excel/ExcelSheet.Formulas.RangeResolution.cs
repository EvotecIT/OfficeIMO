namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private bool TryResolveFormulaRangeReferenceWithError(string token, out ExcelSheet sheet,
            out int r1, out int c1, out int r2, out int c2, out FormulaArgumentValue error) {
            FormulaArgumentValue previousError = _formulaReferenceError;
            _formulaReferenceError = default;
            error = default;
            try {
                bool resolved = TryResolveFormulaRangeReference(token, out sheet, out r1, out c1, out r2, out c2);
                error = _formulaReferenceError;
                return resolved;
            } finally {
                _formulaReferenceError = previousError.IsError ? previousError : error;
            }
        }

        private bool TryResolveFormulaRange(string token, out List<FormulaArgumentValue> values) {
            int remainingCellBudget = MaxResolvedFormulaRangeCells;
            return TryResolveFormulaRange(token, out values, ref remainingCellBudget);
        }

        private bool TryResolveFormulaRange(string token, out List<FormulaArgumentValue> values, ref int remainingCellBudget) {
            return TryResolveFormulaRange(token, out values, ref remainingCellBudget, out _);
        }

        private bool TryResolveFormulaRange(string token, out List<FormulaArgumentValue> values,
            ref int remainingCellBudget, out bool isReference) {
            values = new List<FormulaArgumentValue>();
            isReference = false;
            bool resolved = TryResolveFormulaRangeReferenceWithError(token, out ExcelSheet sheet,
                out int r1, out int c1, out int r2, out int c2, out FormulaArgumentValue referenceError);
            if (!resolved) {
                // A qualified OFFSET failure belongs to this expression. Keep
                // resolving the other inputs before publishing its typed error.
                if (referenceError.IsError) { isReference = true; return true; }
                string expression = token.Trim();
                if (expression.StartsWith("{", StringComparison.Ordinal) || IsSupportedArrayCall(expression)) {
                    isReference = true;
                    if (!TryEvaluateArrayValue(expression, 0, out FormulaArrayValue array)
                        || array.Values.Length > remainingCellBudget) return false;
                    remainingCellBudget -= array.Values.Length;
                    values = array.Values.ToList();
                    return true;
                }
                if (TryResolveDefinedNameArray(expression, 0, out FormulaArrayValue namedArray)) {
                    isReference = true;
                    if (namedArray.Values.Length > remainingCellBudget) return false;
                    remainingCellBudget -= namedArray.Values.Length;
                    values = namedArray.Values.ToList();
                    return true;
                }
                return false;
            }
            isReference = true;

            if (!TryGetFormulaRangeCellCount(r1, c1, r2, c2, out int cellCount) || cellCount > remainingCellBudget) {
                return false;
            }

            remainingCellBudget -= cellCount;
            if (TryReuseCalculatedRange(sheet, r1, c1, r2, c2, out values)) return true;
            values = new List<FormulaArgumentValue>(cellCount);
            FormulaEvaluationDepthFrame? frame = BeginCalculatedRange();
            bool complete = false;
            try {
                for (int row = r1; row <= r2; row++) {
                    for (int column = c1; column <= c2; column++) {
                        values.Add(ResolveFormulaDependency(sheet, row, column));
                    }
                }
                complete = true;
            } finally {
                if (frame != null) EndCalculatedRange(frame, complete, sheet, r1, c1, r2, c2, values);
            }

            return true;
        }

        /// <summary>
        /// Returns a range already resolved in this pass. Its cells cannot change within a pass, so aggregates over a
        /// shared column read it once; the dependency depth it recorded is charged as resolving each cell would.
        /// </summary>
        private bool TryReuseCalculatedRange(ExcelSheet sheet, int r1, int c1, int r2, int c2, out List<FormulaArgumentValue> values) {
            values = null!;
            if (_formulaCalculationContext is not { IsDisposed: false } context
                || _formulaEvaluationStack == null
                || _formulaEvaluationDepthFrames is not { Count: > 0 } frames
                || !context.Ranges.TryGetValue((sheet._worksheetPartValue, r1, c1, r2, c2), out var range)
                || _formulaEvaluationStack.Count + range.Depth > _excelDocument.Calculation.MaximumDependencyDepth) return false;
            frames.Peek().IncludeChild(range.Depth);
            values = new List<FormulaArgumentValue>(range.Values);
            return true;
        }

        /// <summary>Collects the range's dependency outcome in a frame of its own, so it can be kept and later charged again.</summary>
        private FormulaEvaluationDepthFrame? BeginCalculatedRange() {
            if (_formulaCalculationContext is not { IsDisposed: false }
                || _formulaEvaluationCache == null
                || _formulaEvaluationDepthFrames is not { Count: > 0 } frames
                || frames.Peek().DependencyGuardBlocked) return null;
            var frame = new FormulaEvaluationDepthFrame();
            frames.Push(frame);
            return frame;
        }

        private void EndCalculatedRange(FormulaEvaluationDepthFrame frame, bool complete, ExcelSheet sheet,
            int r1, int c1, int r2, int c2, List<FormulaArgumentValue> values) {
            _formulaEvaluationDepthFrames!.Pop();
            FormulaEvaluationDepthFrame parent = _formulaEvaluationDepthFrames.Peek();
            parent.IncludeChild(frame.MaximumChildDepth);
            if (frame.DependencyGuardBlocked) parent.BlockByDependencyGuard();
            if (frame.UsedUnevaluatedFormulaCache) parent.MarkUnevaluatedFormulaCache();
            if (complete && !frame.DependencyGuardBlocked && !frame.UsedUnevaluatedFormulaCache
                && _formulaCalculationContext is { IsDisposed: false } context
                && !values.Exists(static value => value.IsUnresolvedFormula)) {
                context.Ranges[(sheet._worksheetPartValue, r1, c1, r2, c2)] = (values.ToArray(), frame.MaximumChildDepth);
            }
        }

        private static bool TryGetFormulaRangeCellCount(int r1, int c1, int r2, int c2, out int cellCount) {
            cellCount = 0;
            long rowCount = (long)r2 - r1 + 1;
            long columnCount = (long)c2 - c1 + 1;
            long totalCells = rowCount * columnCount;
            if (rowCount <= 0 || columnCount <= 0 || totalCells > MaxResolvedFormulaRangeCells) {
                return false;
            }

            cellCount = (int)totalCells;
            return true;
        }

    }
}
