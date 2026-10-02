using System.Threading;

namespace OfficeIMO.IWork.Internal;

internal static class IWorkNumericDisplayDiagnostics {
    internal static IEnumerable<IWorkDiagnostic> ForTextTables(IEnumerable<IWorkTableCell> cells,
        string sourceLabel, string destinationLabel, CancellationToken cancellationToken) {
        bool formatted = false, omitted = false;
        foreach (IWorkTableCell cell in cells) {
            cancellationToken.ThrowIfCancellationRequested();
            if (cell.NumberFormat == null || cell.ValueKind is not (IWorkCellKind.Number or IWorkCellKind.Duration) || cell.Value == null) continue;
            if (cell.TryGetFormattedNumber(out _, out _)) formatted = true;
            else omitted = true;
        }
        if (formatted) yield return new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
            "IWORK_" + sourceLabel + "_NUMBER_FORMAT_APPROXIMATED",
            destinationLabel + " table cells retain supported invariant numeric and fixed hour/minute duration display as editable text. Automatic precision, currency symbols and locale placement, fraction spacing, duration unit spacing or rounding, and full source appearance can differ. Raw values, formula caches and formats remain available on the source projection.",
            lossKind: global::OfficeIMO.OfficeConversionLossKind.Approximation);
        if (omitted) yield return new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
            "IWORK_" + sourceLabel + "_NUMBER_FORMAT_OMITTED",
            destinationLabel + " retains raw cached text for numeric cells whose display cannot be safely formatted, or preserves their rich text. Their source numeric display formats remain available on the projection.",
            lossKind: global::OfficeIMO.OfficeConversionLossKind.Omission);
    }
}
