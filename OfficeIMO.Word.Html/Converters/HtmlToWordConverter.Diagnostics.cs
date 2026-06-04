namespace OfficeIMO.Word.Html {
    internal partial class HtmlToWordConverter {
        private static void AddDiagnostic(HtmlToWordOptions options, string code, string message, string? source = null, Exception? exception = null, HtmlConversionDiagnosticSeverity severity = HtmlConversionDiagnosticSeverity.Warning) {
            var detail = exception is HtmlConversionLimitException limitException && !string.IsNullOrEmpty(limitException.Detail)
                ? $"{exception.GetType().Name}: {limitException.Detail}"
                : exception is HtmlUnsupportedCssException cssException && !string.IsNullOrEmpty(cssException.Detail)
                    ? $"{exception.GetType().Name}: {cssException.Detail}"
                : exception == null ? null : $"{exception.GetType().Name}: {exception.Message}";
            var diagnostic = new HtmlConversionDiagnostic(code, message, severity, source, detail);
            options.Diagnostics.Add(diagnostic);
            options.DiagnosticHandler?.Invoke(diagnostic);
        }
    }
}
