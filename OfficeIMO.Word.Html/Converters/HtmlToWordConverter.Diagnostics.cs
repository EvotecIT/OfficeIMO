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
            OfficeConversionLossKind? lossKind = null) {
            var detail = exception is HtmlConversionLimitException limitException && !string.IsNullOrEmpty(limitException.Detail)
                ? $"{exception.GetType().Name}: {limitException.Detail}"
                : exception is HtmlUnsupportedCssException cssException && !string.IsNullOrEmpty(cssException.Detail)
                    ? $"{exception.GetType().Name}: {cssException.Detail}"
                : exception == null ? null : $"{exception.GetType().Name}: {exception.Message}";
            OfficeConversionLossKind effectiveLoss = lossKind ?? (severity == HtmlDiagnosticSeverity.Error
                ? OfficeConversionLossKind.Failure
                : severity == HtmlDiagnosticSeverity.Info
                    ? OfficeConversionLossKind.None
                    : OfficeConversionLossKind.Omission);
            options.ConversionReport.Add("OfficeIMO.Word.Html", code, message, severity, source, detail, effectiveLoss);
        }
    }
}
