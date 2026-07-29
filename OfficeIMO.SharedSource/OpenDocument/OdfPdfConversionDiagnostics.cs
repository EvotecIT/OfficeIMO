using System;
using System.Collections.Generic;
using System.Globalization;
using OfficeIMO.OpenDocument;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.OpenDocument.Internal;

internal static class OdfPdfConversionDiagnostics {
    internal static PdfCore.PdfDocumentConversionResult Attach(
        PdfCore.PdfDocumentConversionResult result,
        OdfConversionReport report,
        string converterName) {
        if (result == null) throw new ArgumentNullException(nameof(result));
        if (report == null) throw new ArgumentNullException(nameof(report));
        if (string.IsNullOrWhiteSpace(converterName)) {
            throw new ArgumentException("Converter name cannot be empty.", nameof(converterName));
        }

        var warnings = new List<PdfCore.PdfConversionWarning>(report.Mappings.Count);
        foreach (OdfConversionMapping mapping in report.Mappings) {
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
            warnings.Add(new PdfCore.PdfConversionWarning(
                converterName,
                "ODF_" + mapping.Status.ToString().ToUpperInvariant(),
                report.SourceFormat + "->" + report.TargetFormat + ":" + mapping.Feature,
                message,
                severity,
                details: details));
        }

        return result.WithAdditionalWarnings(warnings);
    }
}
