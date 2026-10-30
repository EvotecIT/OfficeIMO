namespace OfficeIMO.Pdf;

public sealed partial class PdfCreateOptions {
    private bool _suppressDiagnostics;
    private PdfConversionReport? DiagnosticsReport => _suppressDiagnostics
        ? null : _diagnosticsReport ?? PdfOutputDiagnosticScope.CurrentReport;


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
