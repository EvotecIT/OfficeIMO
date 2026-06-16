namespace OfficeIMO.Html;

public sealed partial class HtmlToRtfOptions {
    internal void AddDiagnostic(string code, string message, string? source = null, Exception? exception = null, HtmlRtfConversionDiagnosticSeverity severity = HtmlRtfConversionDiagnosticSeverity.Warning) {
        string? detail = exception is HtmlRtfConversionLimitException limitException && !string.IsNullOrEmpty(limitException.Detail)
            ? exception.GetType().Name + ": " + limitException.Detail
            : exception == null ? null : exception.GetType().Name + ": " + exception.Message;

        var diagnostic = new HtmlRtfConversionDiagnostic(code, message, severity, source, detail);
        Diagnostics.Add(diagnostic);
        DiagnosticHandler?.Invoke(diagnostic);
    }
}
