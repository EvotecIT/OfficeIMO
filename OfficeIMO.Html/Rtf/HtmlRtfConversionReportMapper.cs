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
            diagnostic.Action,
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
        RtfConversionAction action = diagnostic.Action;
        HtmlConversionLossKind lossKind = diagnostic.Severity == HtmlRtfConversionDiagnosticSeverity.Error
            ? HtmlConversionLossKind.Failure
            : action == RtfConversionAction.Substituted || action == RtfConversionAction.Flattened
                ? HtmlConversionLossKind.Approximation
                : action == RtfConversionAction.Omitted || action == RtfConversionAction.Blocked
                    ? HtmlConversionLossKind.Omission
                    : HtmlConversionLossKind.None;
        report.Add("OfficeIMO.Html.Rtf", diagnostic.Code, diagnostic.Message, severity, diagnostic.Source, diagnostic.Detail, lossKind);
    }

}
