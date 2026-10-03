namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        // One context per workbook calculation, or per standalone sheet calculation.
        // Referenced sheets already borrow these objects when resolving dependencies.
        internal sealed class FormulaCalculationContext {
            internal Dictionary<string, FormulaArgumentValue> Cache { get; } = new(StringComparer.OrdinalIgnoreCase);
            internal Dictionary<string, int> DepthCache { get; } = new(StringComparer.OrdinalIgnoreCase);
            internal HashSet<string> Stack { get; } = new(StringComparer.OrdinalIgnoreCase);
            internal Stack<FormulaEvaluationDepthFrame> DepthFrames { get; } = new();
            internal FormulaEvaluationGuardState GuardState { get; } = new();
        }

        // Each expression owns reference failures. A nested expression returns its
        // typed error to IFERROR/ISERROR rather than leaking it into the parent scope.
        private FormulaArgumentValue _formulaReferenceError;

        private bool TryEvaluateFormulaValue(string formula, out FormulaArgumentValue result, bool allowScalarExpression = true) {
            result = default;
            if (!HasSufficientFormulaExecutionStack() || _scalarFormulaEvaluationDepth >= 128) return false;
            _scalarFormulaEvaluationDepth++;
            FormulaArgumentValue previousError = _formulaReferenceError;
            _formulaReferenceError = default;
            try {
                bool evaluated;
                if (allowScalarExpression) {
                    if (string.IsNullOrWhiteSpace(formula) || formula.Length > MaxSupportedFormulaLength) return false;
                    evaluated = TryEvaluateScalarExpression(NormalizeSupportedFunctionPrefix(formula), out result);
                } else {
                    evaluated = TryEvaluateFormulaValueCore(formula, out result);
                }
                if (_formulaReferenceError.IsError) { result = _formulaReferenceError; return true; }
                return evaluated;
            } finally {
                _formulaReferenceError = previousError;
                _scalarFormulaEvaluationDepth--;
            }
        }

    }
}
