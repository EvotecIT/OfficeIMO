namespace OfficeIMO.Html.Rtf;

public sealed partial class RtfHtmlReadOptions {
    internal void AddDiagnostic(string code, string message, string? source = null, Exception? exception = null, RtfHtmlConversionDiagnosticSeverity severity = RtfHtmlConversionDiagnosticSeverity.Warning) {
        string? detail = exception is RtfHtmlConversionLimitException limitException && !string.IsNullOrEmpty(limitException.Detail)
            ? exception.GetType().Name + ": " + limitException.Detail
            : exception == null ? null : exception.GetType().Name + ": " + exception.Message;

        var diagnostic = new RtfHtmlConversionDiagnostic(code, message, severity, source, detail);
        Diagnostics.Add(diagnostic);
        DiagnosticHandler?.Invoke(diagnostic);
    }
}
