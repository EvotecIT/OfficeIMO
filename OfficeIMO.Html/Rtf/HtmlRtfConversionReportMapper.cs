namespace OfficeIMO.Html;

internal static class HtmlRtfConversionReportMapper {
    public static void Add(RtfConversionReport report, HtmlRtfConversionDiagnostic diagnostic) {
        RtfConversionSeverity severity = diagnostic.Severity == HtmlRtfConversionDiagnosticSeverity.Error
            ? RtfConversionSeverity.Error
            : diagnostic.Severity == HtmlRtfConversionDiagnosticSeverity.Warning
                ? RtfConversionSeverity.Warning
                : RtfConversionSeverity.Information;
        report.Add(
            severity,
            diagnostic.Code,
            diagnostic.Message,
            GetAction(diagnostic),
            sourcePath: diagnostic.Source,
            feature: diagnostic.Source,
            detail: diagnostic.Detail);
    }

    private static RtfConversionAction GetAction(HtmlRtfConversionDiagnostic diagnostic) {
        if (diagnostic.Severity == HtmlRtfConversionDiagnosticSeverity.Info) return RtfConversionAction.Preserved;
        if (diagnostic.Code.IndexOf("Rejected", StringComparison.OrdinalIgnoreCase) >= 0 ||
            diagnostic.Code.IndexOf("Blocked", StringComparison.OrdinalIgnoreCase) >= 0 ||
            diagnostic.Code.IndexOf("LimitExceeded", StringComparison.OrdinalIgnoreCase) >= 0) {
            return RtfConversionAction.Blocked;
        }

        if (diagnostic.Code.IndexOf("Fallback", StringComparison.OrdinalIgnoreCase) >= 0 ||
            diagnostic.Code.IndexOf("Substitut", StringComparison.OrdinalIgnoreCase) >= 0) {
            return RtfConversionAction.Substituted;
        }

        return RtfConversionAction.Omitted;
    }
}
