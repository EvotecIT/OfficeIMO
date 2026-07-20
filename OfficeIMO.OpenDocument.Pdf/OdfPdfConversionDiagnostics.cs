using System.Globalization;
using System.Text;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.OpenDocument.Pdf;

internal static class OdfPdfConversionDiagnostics {
    internal static PdfCore.PdfDocumentConversionResult Attach(
        PdfCore.PdfDocumentConversionResult result,
        OdfConversionReport report) {
        if (result == null) throw new ArgumentNullException(nameof(result));
        if (report == null) throw new ArgumentNullException(nameof(report));

        var warnings = new List<PdfCore.PdfConversionWarning>(report.Mappings.Count);
        foreach (OdfConversionMapping mapping in report.Mappings) {
            warnings.Add(ToWarning(report, mapping));
        }

        return result.WithAdditionalWarnings(warnings);
    }

    private static PdfCore.PdfConversionWarning ToWarning(
        OdfConversionReport report,
        OdfConversionMapping mapping) {
        PdfCore.PdfConversionWarningSeverity severity =
            mapping.Status == OdfConversionMappingStatus.Converted
                ? PdfCore.PdfConversionWarningSeverity.Information
                : PdfCore.PdfConversionWarningSeverity.Warning;
        string message = string.IsNullOrWhiteSpace(mapping.Message)
            ? string.Format(
                CultureInfo.InvariantCulture,
                "{0} {1} item(s) were {2} while projecting {3} to {4} before PDF layout.",
                mapping.Count,
                mapping.Feature,
                mapping.Status.ToString().ToLowerInvariant(),
                report.SourceFormat,
                report.TargetFormat)
            : mapping.Message!;
        var details = new Dictionary<string, string> {
            ["stage"] = "open-document-projection",
            ["sourceFormat"] = report.SourceFormat,
            ["targetFormat"] = report.TargetFormat,
            ["feature"] = mapping.Feature,
            ["status"] = mapping.Status.ToString(),
            ["count"] = mapping.Count.ToString(CultureInfo.InvariantCulture)
        };
        return new PdfCore.PdfConversionWarning(
            "OfficeIMO.OpenDocument.Pdf",
            "ODF_" + NormalizeToken(mapping.Status.ToString()),
            report.SourceFormat + "->" + report.TargetFormat + ":" + mapping.Feature,
            message,
            severity,
            details: details);
    }

    private static string NormalizeToken(string value) {
        var builder = new StringBuilder(value.Length);
        bool previousUnderscore = false;
        foreach (char character in value) {
            char normalized = char.IsLetterOrDigit(character)
                ? char.ToUpperInvariant(character)
                : '_';
            if (normalized == '_' && previousUnderscore) continue;
            builder.Append(normalized);
            previousUnderscore = normalized == '_';
        }

        return builder.ToString().Trim('_');
    }
}
