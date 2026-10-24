namespace OfficeIMO.Pdf;

public sealed partial class PdfOptions {
    // Clones share report identity, while independent conversions own separate buffers.
    internal PdfLayoutDiagnosticBuffer BeginLayoutDiagnosticBuffer() => new(_diagnosticsReport);

    internal void AddUnassessedLayoutWarning(string code, string source, string message) {
        if (_diagnosticsReport == null) return;
        if (PdfLayoutDiagnosticBuffer.TryCapture(_diagnosticsReport, () => AddUnassessedLayoutWarning(code, source, message))) return;
        string key = code + "|" + source + "|" + message;
        if (!(_reportedLayoutDiagnostics ??= new HashSet<string>()).Add(key)) return;
        _diagnosticsReport.Add(new PdfConversionWarning(_diagnosticsConverter, code, source, message,
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
        if (_diagnosticsReport == null) {
            return;
        }

        if (PdfLayoutDiagnosticBuffer.TryCapture(_diagnosticsReport, () => AddLayoutDiagnostic(code, source, message,
                kind, severity, x, y, width, height))) return;

        string key = code + "|" + source + "|" + message;
        if (!(_reportedLayoutDiagnostics ??= new HashSet<string>()).Add(key)) {
            return;
        }

        var layoutDiagnostic = new PdfLayoutDiagnostic(kind, source, message, x, y, width, height);
        _diagnosticsReport.Add(new PdfConversionWarning(
            _diagnosticsConverter,
            code,
            source,
            message,
            severity,
            layoutDiagnostic));
    }
}
