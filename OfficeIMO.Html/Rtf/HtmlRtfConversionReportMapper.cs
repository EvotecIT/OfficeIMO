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

    public static void Add(HtmlDiagnosticReport report, HtmlRtfConversionDiagnostic diagnostic) {
        HtmlDiagnosticSeverity severity = diagnostic.Severity == HtmlRtfConversionDiagnosticSeverity.Error
            ? HtmlDiagnosticSeverity.Error
            : diagnostic.Severity == HtmlRtfConversionDiagnosticSeverity.Warning
                ? HtmlDiagnosticSeverity.Warning
                : HtmlDiagnosticSeverity.Info;
        RtfConversionAction action = GetAction(diagnostic);
        HtmlConversionLossKind lossKind = diagnostic.Severity == HtmlRtfConversionDiagnosticSeverity.Error
            ? HtmlConversionLossKind.Failure
            : action == RtfConversionAction.Substituted
                ? HtmlConversionLossKind.Approximation
                : action == RtfConversionAction.Omitted || action == RtfConversionAction.Blocked
                    ? HtmlConversionLossKind.Omission
                    : HtmlConversionLossKind.None;
        report.Add("OfficeIMO.Html.Rtf", diagnostic.Code, diagnostic.Message, severity, diagnostic.Source, diagnostic.Detail, lossKind);
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
