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
            values = new List<FormulaArgumentValue>(cellCount);
            for (int row = r1; row <= r2; row++) {
                for (int column = c1; column <= c2; column++) {
                    values.Add(ResolveFormulaDependency(sheet, row, column));
                }
            }

            return true;
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
