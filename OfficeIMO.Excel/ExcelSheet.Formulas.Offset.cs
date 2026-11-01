namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private int _offsetFormulaReferenceDepth;

        private bool TryEvaluateOffsetValue(string args, out FormulaArgumentValue result) {
            result = default;
            int? currentRow = _formulaEvaluationCellReference != null
                && TryParseCellReference(_formulaEvaluationCellReference, out int row, out _) ? row : null;
            if (!TryResolveOffsetRange(args, currentRow, out ExcelSheet sheet, out int r1, out int c1, out int r2, out int c2, out result)) return false;
            if (result.IsError) return true;
            // Multi-cell OFFSET is a reference argument, not a scalar or a dynamic spill in this subset.
            if (r1 != r2 || c1 != c2) return false;
            result = ResolveFormulaDependency(sheet, r1, c1);
            return !result.IsUnresolvedFormula;
        }

        private bool TryResolveOffsetRange(string args, int? currentRow, out ExcelSheet sheet,
            out int r1, out int c1, out int r2, out int c2, out FormulaArgumentValue error) {
            sheet = this;
            r1 = c1 = r2 = c2 = 0;
            error = default;
            if (_offsetFormulaReferenceDepth >= 32 || !HasSufficientFormulaExecutionStack()) return false;
            _offsetFormulaReferenceDepth++;
            FormulaArgumentValue previousError = _formulaReferenceError;
            _formulaReferenceError = default;
            try {
                IReadOnlyList<string> tokens = SplitFormulaArguments(args, preserveEmpty: true);
                if (tokens.Count < 3 || tokens.Count > 5 || string.IsNullOrWhiteSpace(tokens[0])
                    || string.IsNullOrWhiteSpace(tokens[1]) || string.IsNullOrWhiteSpace(tokens[2])) return false;
                bool baseResolved = TryResolveFormulaRangeReference(tokens[0], currentRow, out sheet, out int baseRow, out int baseColumn,
                    out int baseEndRow, out int baseEndColumn);
                if (!baseResolved && !_formulaReferenceError.IsError) return false;
                error = _formulaReferenceError;
                double[] parameters = { 0d, 0d, (double)baseEndRow - baseRow + 1d, (double)baseEndColumn - baseColumn + 1d };
                for (int index = 1; index < tokens.Count; index++) {
                    if (index >= 3 && string.IsNullOrWhiteSpace(tokens[index])) continue;
                    if (!TryResolveFormulaValueArgument(tokens[index], out FormulaArgumentValue value)) return false;
                    if (value.IsError) { if (!error.IsError) error = value; continue; }
                    if (!IsFiniteFormulaNumber(value)) return false;
                    parameters[index - 1] = Math.Truncate(value.Number!.Value);
                }
                if (error.IsError) return true;
                double top = baseRow + parameters[0], left = baseColumn + parameters[1];
                double bottom = top + parameters[2] - 1d, right = left + parameters[3] - 1d;
                if (parameters[2] <= 0d || parameters[3] <= 0d || top < 1d || left < 1d
                    || top > A1.MaxRows || left > A1.MaxColumns || bottom > A1.MaxRows || right > A1.MaxColumns) {
                    error = FormulaArgumentValue.Error("#REF!");
                    return true;
                }
                r1 = (int)top; c1 = (int)left; r2 = (int)bottom; c2 = (int)right;
                return true;
            } finally {
                _formulaReferenceError = previousError;
                _offsetFormulaReferenceDepth--;
            }
        }
    }
}
