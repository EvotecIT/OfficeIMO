using OfficeIMO.Html;

namespace OfficeIMO.Word.Html {
    internal partial class HtmlToWordConverter {
        private static void AddDiagnostic(
            HtmlToWordOptions options,
            string code,
            string message,
            string? source = null,
            Exception? exception = null,
            HtmlDiagnosticSeverity severity = HtmlDiagnosticSeverity.Warning,
            HtmlConversionLossKind? lossKind = null) {
            var detail = exception is HtmlConversionLimitException limitException && !string.IsNullOrEmpty(limitException.Detail)
                ? $"{exception.GetType().Name}: {limitException.Detail}"
                : exception is HtmlUnsupportedCssException cssException && !string.IsNullOrEmpty(cssException.Detail)
                    ? $"{exception.GetType().Name}: {cssException.Detail}"
                : exception == null ? null : $"{exception.GetType().Name}: {exception.Message}";
            HtmlConversionLossKind effectiveLoss = lossKind ?? (severity == HtmlDiagnosticSeverity.Error
                ? HtmlConversionLossKind.Failure
                : severity == HtmlDiagnosticSeverity.Info
                    ? HtmlConversionLossKind.None
                    : HtmlConversionLossKind.Omission);
            options.ConversionReport.Add("OfficeIMO.Word.Html", code, message, severity, source, detail, effectiveLoss);
        }
    }
}
