namespace OfficeIMO.Pdf;

public sealed partial class PdfCreateOptions {
    private bool _suppressDiagnostics;
    private PdfConversionReport? DiagnosticsReport => _suppressDiagnostics
        ? null : _diagnosticsReport ?? PdfOutputDiagnosticScope.CurrentReport;


    /// <summary>
    /// When set, clipped-cell findings of tables that name their source cells are collected here and reported once
    /// per contiguous range when the layout completes, instead of once per cell.
    /// </summary>
    internal PdfTableCellDiagnosticSummary? TableCellDiagnosticSummary { get; set; }

    internal void AddTableCellDiagnostic(PdfTableCellDiagnosticSummary summary, string code, PdfTableCellDiagnosticAddress address,
        string message, PdfLayoutDiagnosticKind kind, PdfConversionWarningSeverity severity) {
        if (DiagnosticsReport == null) return;
        if (PdfLayoutDiagnosticBuffer.TryCapture(DiagnosticsReport, () => AddTableCellDiagnostic(summary, code, address, message, kind, severity))) return;
        summary.Record(code, address, message, kind, severity);
    }

    /// <summary>Reports the summarized ranges of an accepted layout.</summary>
    internal void CompleteTableCellDiagnosticSummary() {
        PdfTableCellDiagnosticSummary? summary = TableCellDiagnosticSummary;
        if (summary == null || DiagnosticsReport == null) return;
        foreach (PdfTableCellDiagnosticRange range in summary.Complete())
            AddLayoutDiagnostic(range.Code, summary.FormatSource(range), summary.FormatMessage(range), range.Kind, range.Severity);
    }

    // Clones share report identity, while independent conversions own separate buffers.
    internal PdfLayoutDiagnosticBuffer BeginLayoutDiagnosticBuffer() => new(DiagnosticsReport);

    internal void AddUnassessedLayoutWarning(string code, string source, string message) {
        if (DiagnosticsReport == null) return;
        if (PdfLayoutDiagnosticBuffer.TryCapture(DiagnosticsReport, () => AddUnassessedLayoutWarning(code, source, message))) return;
        string key = code + "|" + source + "|" + message;
        if (!PdfOutputDiagnosticScope.Admit("layout|" + key, ref _reportedLayoutDiagnostics)) return;
        DiagnosticsReport.Add(new PdfConversionWarning(_diagnosticsConverter, code, source, message,
            PdfConversionWarningSeverity.Information, OfficeConversionLossKind.Unassessed));
    }

    internal void AddLayoutDiagnostic(
        string code,
        string source,
        string message,
        PdfLayoutDiagnosticKind kind,
        PdfConversionWarningSeverity severity = PdfConversionWarningSeverity.Warning,
        double? x = null,
        double? y = null,
        double? width = null,
        double? height = null) {
        if (DiagnosticsReport == null) {
            return;
        }

        if (PdfLayoutDiagnosticBuffer.TryCapture(DiagnosticsReport, () => AddLayoutDiagnostic(code, source, message,
                kind, severity, x, y, width, height))) return;

        string key = code + "|" + source + "|" + message;
        if (!PdfOutputDiagnosticScope.Admit("layout|" + key, ref _reportedLayoutDiagnostics)) {
            return;
        }

        var layoutDiagnostic = new PdfLayoutDiagnostic(kind, source, message, (x.HasValue ? global::OfficeIMO.OfficeLength.Points(x.Value) : (global::OfficeIMO.OfficeLength?)null), (y.HasValue ? global::OfficeIMO.OfficeLength.Points(y.Value) : (global::OfficeIMO.OfficeLength?)null), (width.HasValue ? global::OfficeIMO.OfficeLength.Points(width.Value) : (global::OfficeIMO.OfficeLength?)null), (height.HasValue ? global::OfficeIMO.OfficeLength.Points(height.Value) : (global::OfficeIMO.OfficeLength?)null));
        DiagnosticsReport.Add(new PdfConversionWarning(
            _diagnosticsConverter,
            code,
            source,
            message,
            severity,
            layoutDiagnostic));
    }
}
