namespace OfficeIMO.Html;

public sealed partial class HtmlToRtfOptions {
    internal void AddDiagnostic(string code, string message, string? source = null, Exception? exception = null, HtmlRtfConversionDiagnosticSeverity severity = HtmlRtfConversionDiagnosticSeverity.Warning, RtfConversionAction? action = null) {
        string? detail = exception is HtmlRtfConversionLimitException limitException && !string.IsNullOrEmpty(limitException.Detail)
            ? exception.GetType().Name + ": " + limitException.Detail
            : exception == null ? null : exception.GetType().Name + ": " + exception.Message;

        var diagnostic = new HtmlRtfConversionDiagnostic(code, message, severity, source, detail, action);
        Diagnostics.Add(diagnostic);
        HtmlRtfConversionReportMapper.Add(ConversionReport, diagnostic);
        HtmlRtfConversionReportMapper.Add(HtmlDiagnostics, diagnostic);
    }

    internal void AddDiagnostic(HtmlRtfConversionDiagnostic diagnostic) {
        if (diagnostic == null) throw new ArgumentNullException(nameof(diagnostic));
        Diagnostics.Add(diagnostic);
        HtmlRtfConversionReportMapper.Add(ConversionReport, diagnostic);
        HtmlRtfConversionReportMapper.Add(HtmlDiagnostics, diagnostic);
    }
}
