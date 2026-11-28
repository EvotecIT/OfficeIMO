namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Evaluates formula cells in a range that carry no cached result, without writing
        /// anything back to the worksheet. Renderers use this to show what a spreadsheet
        /// application would show after recalculation instead of the formula text.
        /// </summary>
        /// <returns>Calculated display values keyed by one-based row and column, plus the
        /// references of formulas the lightweight engine could not evaluate.</returns>
        internal (IReadOnlyDictionary<(int Row, int Column), object?> Values, IReadOnlyList<string> Unsupported) EvaluateUncachedFormulaDisplayValues(
            int firstRow, int firstColumn, int lastRow, int lastColumn) {
            var values = new Dictionary<(int Row, int Column), object?>();
            var unsupported = new List<string>();
            EvaluateUncachedFormulaCells(firstRow, firstColumn, lastRow, lastColumn, writeCache: false, (row, column, reference, _, result) => {
                if (result == null) {
                    values[(row, column)] = null;
                    unsupported.Add(reference);
                    return;
                }

                FormulaArgumentValue value = result.Value;
                values[(row, column)] = value.IsError ? value.Text
                    : value.IsBoolean ? value.Number.GetValueOrDefault() != 0D
                    : value.Number.HasValue ? value.Number.Value
                    : value.Text;
            });
            return (values, unsupported);
        }

        /// <summary>Calculates uncached formulas in a range and stores their results as cached values.</summary>
        private void CacheUncachedFormulaResults(int firstRow, int firstColumn, int lastRow, int lastColumn) =>
            EvaluateUncachedFormulaCells(firstRow, firstColumn, lastRow, lastColumn, writeCache: true, (_, _, _, _, _) => { });

        private void EvaluateUncachedFormulaCells(
            int firstRow,
            int firstColumn,
            int lastRow,
            int lastColumn,
            bool writeCache,
            Action<int, int, string, ExcelFormulaCell, FormulaArgumentValue?> visit) {
            MaterializePendingDirectCellValues();
            Locking.ExecuteWrite(_excelDocument.EnsureLock(), () => {
                var candidates = new List<(ExcelFormulaCell Cell, int Row, int Column, string Reference)>();
                foreach (ExcelFormulaCell cell in EnumerateFormulaCells()) {
                    if (cell.HasValue || cell.HasInlineString) continue;
                    string? reference = cell.Reference;
                    if (reference == null || !TryParseCellReference(reference, out int row, out int column) ||
                        row < firstRow || row > lastRow || column < firstColumn || column > lastColumn) {
                        continue;
                    }

                    candidates.Add((cell, row, column, reference));
                }

                if (candidates.Count == 0) return;

                var context = new FormulaCalculationContext();
                var previousCache = _formulaEvaluationCache;
                var previousDepthCache = _formulaEvaluationDepthCache;
                var previousStack = _formulaEvaluationStack;
                var previousDepthFrames = _formulaEvaluationDepthFrames;
                var previousGuardState = _formulaEvaluationGuardState;
                var previousSharedDefinitions = _formulaEvaluationSharedDefinitions;
                var previousSharedDefinitionsBySheet = _formulaEvaluationSharedDefinitionsBySheet;
                _formulaEvaluationCache = context.Cache;
                _formulaEvaluationDepthCache = context.DepthCache;
                _formulaEvaluationStack = context.Stack;
                _formulaEvaluationDepthFrames = context.DepthFrames;
                _formulaEvaluationGuardState = context.GuardState;
                _formulaEvaluationSharedDefinitions = BuildSharedFormulaDefinitions();
                _formulaEvaluationSharedDefinitionsBySheet = new Dictionary<string, IReadOnlyDictionary<uint, SharedFormulaDefinition>>(
                    StringComparer.OrdinalIgnoreCase) {
                    [Name] = _formulaEvaluationSharedDefinitions
                };
                bool changed = false;
                try {
                    foreach ((ExcelFormulaCell cell, int row, int column, string reference) in candidates) {
                        _formulaEvaluationGuardState.DependencyGuardBlocked = false;
                        if (!TryEvaluateFormulaCellValue(cell, out FormulaArgumentValue result) ||
                            result.IsUnresolvedFormula || !result.HasValue) {
                            visit(row, column, reference, cell, null);
                            continue;
                        }

                        if (writeCache) {
                            SetFormulaCachedValue(cell, result);
                            cell.ClearCalculateCell();
                            changed = true;
                        }

                        visit(row, column, reference, cell, result);
                    }
                } finally {
                    _formulaEvaluationCache = previousCache;
                    _formulaEvaluationDepthCache = previousDepthCache;
                    _formulaEvaluationStack = previousStack;
                    _formulaEvaluationDepthFrames = previousDepthFrames;
                    _formulaEvaluationGuardState = previousGuardState;
                    _formulaEvaluationSharedDefinitions = previousSharedDefinitions;
                    _formulaEvaluationSharedDefinitionsBySheet = previousSharedDefinitionsBySheet;
                }

                if (changed) {
                    _hasWorksheetMutations = true;
                    MarkRequiresSavePreparation();
                    ClearCellTextSharedStringCache();
                }
            });
        }
    }
}
