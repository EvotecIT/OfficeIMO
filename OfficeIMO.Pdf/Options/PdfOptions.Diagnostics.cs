namespace OfficeIMO.Pdf;

public sealed partial class PdfOptions {
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
