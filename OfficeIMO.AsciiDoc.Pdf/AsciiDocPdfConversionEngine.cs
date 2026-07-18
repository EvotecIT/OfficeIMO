using System;
using System.Collections.Generic;
using OfficeIMO.AsciiDoc;
using OfficeIMO.AsciiDoc.Markdown;
using OfficeIMO.Markdown.Pdf;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.AsciiDoc.Pdf;

internal static class AsciiDocPdfConversionEngine {
    internal static PdfCore.PdfDocumentConversionResult Convert(AsciiDocDocument document, AsciiDocPdfSaveOptions? options) {
        if (document == null) throw new ArgumentNullException(nameof(document));

        AsciiDocPdfSaveOptions operation = (options ?? new AsciiDocPdfSaveOptions()).CloneForConversion();
        AsciiDocToMarkdownResult projection = document.ToMarkdownDocumentResult(operation.ProjectionOptions);
        PdfCore.PdfDocumentConversionResult result = projection.Value.ToPdfDocumentResult(operation.PdfOptions);
        return result.WithAdditionalWarnings(ToPdfWarnings(document, projection));
    }

    private static IEnumerable<PdfCore.PdfConversionWarning> ToPdfWarnings(
        AsciiDocDocument document,
        AsciiDocToMarkdownResult projection) {
        foreach (AsciiDocDiagnostic diagnostic in document.Diagnostics) {
            yield return new PdfCore.PdfConversionWarning(
                "OfficeIMO.AsciiDoc.Pdf",
                diagnostic.Code,
                "parser @ " + diagnostic.Span,
                diagnostic.Message,
                ToPdfSeverity(diagnostic.Severity),
                details: new Dictionary<string, string> {
                    ["stage"] = "parse",
                    ["sourceSpan"] = diagnostic.Span.ToString()
                });
        }

        foreach (AsciiDocMarkdownConversionDiagnostic diagnostic in projection.Report.Diagnostics) {
            yield return new PdfCore.PdfConversionWarning(
                "OfficeIMO.AsciiDoc.Pdf",
                diagnostic.Code,
                diagnostic.Feature + " @ " + diagnostic.SourceSpan,
                diagnostic.Message,
                diagnostic.Outcome == AsciiDocMarkdownConversionOutcome.Converted
                    ? PdfCore.PdfConversionWarningSeverity.Information
                    : PdfCore.PdfConversionWarningSeverity.Warning,
                details: new Dictionary<string, string> {
                    ["stage"] = "semantic-projection",
                    ["feature"] = diagnostic.Feature,
                    ["outcome"] = diagnostic.Outcome.ToString(),
                    ["sourceSpan"] = diagnostic.SourceSpan.ToString()
                });
        }
    }

    private static PdfCore.PdfConversionWarningSeverity ToPdfSeverity(AsciiDocDiagnosticSeverity severity) => severity switch {
        AsciiDocDiagnosticSeverity.Error => PdfCore.PdfConversionWarningSeverity.Error,
        AsciiDocDiagnosticSeverity.Warning => PdfCore.PdfConversionWarningSeverity.Warning,
        _ => PdfCore.PdfConversionWarningSeverity.Information
    };
}
