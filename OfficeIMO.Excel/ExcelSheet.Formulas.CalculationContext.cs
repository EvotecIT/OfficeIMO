namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        // One context per workbook calculation, or per standalone sheet calculation.
        // Referenced sheets already borrow these objects when resolving dependencies.
        internal sealed class FormulaCalculationContext : IDisposable {
            internal Dictionary<string, FormulaArgumentValue> Cache { get; } = new(StringComparer.OrdinalIgnoreCase);
            internal Dictionary<string, int> DepthCache { get; } = new(StringComparer.OrdinalIgnoreCase);
            internal HashSet<string> Stack { get; } = new(StringComparer.OrdinalIgnoreCase);
            internal Stack<FormulaEvaluationDepthFrame> DepthFrames { get; } = new();
            internal FormulaEvaluationGuardState GuardState { get; } = new();
            private SharedStringCache? _sharedStrings;
            private long _sharedStringMutationVersion;
            internal bool IsDisposed { get; private set; }

            // Cross-sheet wrappers borrow this snapshot only while the pass owns it.
            internal SharedStringCache GetSharedStrings(DocumentFormat.OpenXml.Packaging.SpreadsheetDocument document, long mutationVersion) {
                if (_sharedStrings == null || _sharedStringMutationVersion != mutationVersion) {
                    _sharedStrings?.Dispose();
                    _sharedStrings = SharedStringCache.Build(document);
                    _sharedStringMutationVersion = mutationVersion;
                }
                return _sharedStrings;
            }

            public void Dispose() {
                _sharedStrings?.Dispose();
                _sharedStrings = null;
                IsDisposed = true;
            }
        }

        private FormulaCalculationContext? _formulaCalculationContext;

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
                // A reference error is a result only after the owning expression
                // qualifies every required input. It must never turn an unsupported
                // expression into an error that IFERROR can silently replace.
                if (evaluated && _formulaReferenceError.IsError) { result = _formulaReferenceError; return true; }
                return evaluated;
            } finally {
                _formulaReferenceError = previousError;
                _scalarFormulaEvaluationDepth--;
            }
        }

    }
}
