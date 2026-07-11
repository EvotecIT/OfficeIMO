using System;
using OfficeIMO.Rtf;

namespace OfficeIMO.Rtf.Markdown;

internal static class RtfMarkdownConversionReportMapper {
    public static void Add(RtfConversionReport report, RtfMarkdownConversionDiagnostic diagnostic) {
        RtfConversionSeverity severity = diagnostic.Severity == RtfMarkdownDiagnosticSeverity.Error
            ? RtfConversionSeverity.Error
            : diagnostic.Severity == RtfMarkdownDiagnosticSeverity.Warning
                ? RtfConversionSeverity.Warning
                : RtfConversionSeverity.Information;
        report.Add(
            severity,
            diagnostic.Code,
            diagnostic.Message,
            GetAction(diagnostic),
            sourcePath: diagnostic.Source,
            feature: diagnostic.Source);
    }

    private static RtfConversionAction GetAction(RtfMarkdownConversionDiagnostic diagnostic) {
        string message = diagnostic.Message;
        if (message.IndexOf("omitted", StringComparison.OrdinalIgnoreCase) >= 0 ||
            message.IndexOf("unsupported", StringComparison.OrdinalIgnoreCase) >= 0) {
            return RtfConversionAction.Omitted;
        }

        if (message.IndexOf("flattened", StringComparison.OrdinalIgnoreCase) >= 0 ||
            message.IndexOf("fallback", StringComparison.OrdinalIgnoreCase) >= 0 ||
            message.IndexOf("represented", StringComparison.OrdinalIgnoreCase) >= 0) {
            return RtfConversionAction.Flattened;
        }

        return RtfConversionAction.Preserved;
    }
}
