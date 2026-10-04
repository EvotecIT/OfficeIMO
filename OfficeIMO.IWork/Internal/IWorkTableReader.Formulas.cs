namespace OfficeIMO.IWork.Internal;

internal static partial class IWorkTableReader {
    internal static void AssessFormulas(IReadOnlyList<IWorkTableCell> cells, string name, IWorkArchiveRecord model,
        List<IWorkDiagnostic> diagnostics, ref bool supportsEditableReconstruction) {
        int incompleteCachedFormulaCount = cells.Count(cell => cell.Kind == IWorkCellKind.Formula
            && !cell.FormulaIsComplete && cell.Value != null);
        if (incompleteCachedFormulaCount > 0) {
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning, "IWORK_TABLE_FORMULA_PARTIAL",
                $"{incompleteCachedFormulaCount} formulas in table '{name}' retain typed cached values because their expressions were not reconstructed completely.",
                model.EntryPath, model.Identifier));
        }
        int incompleteFormulaCacheCount = cells.Count(cell => cell.Kind == IWorkCellKind.Formula
            && !cell.CachedValueIsComplete && cell.Value != null);
        if (incompleteFormulaCacheCount > 0) {
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_TABLE_FORMULA_CACHE_PARTIAL",
                $"{incompleteFormulaCacheCount} formula cached values in table '{name}' are partial; only complete expressions can be reconstructed as editable formulas.",
                model.EntryPath, model.Identifier));
            if (cells.Any(cell => cell.Kind == IWorkCellKind.Formula
                && !cell.CachedValueIsComplete && !cell.FormulaIsComplete)) {
                supportsEditableReconstruction = false;
            }
        }
        int incompleteUncachedFormulaCount = cells.Count(cell => cell.Kind == IWorkCellKind.Formula
            && (!cell.FormulaIsComplete || !cell.CachedValueIsComplete) && cell.Value == null);
        if (incompleteUncachedFormulaCount > 0) {
            supportsEditableReconstruction = false;
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_TABLE_FORMULA_UNSUPPORTED",
                $"{incompleteUncachedFormulaCount} formulas in table '{name}' have neither a complete expression nor a cached value; editable reconstruction is incomplete.",
                model.EntryPath, model.Identifier));
        }
    }
}
