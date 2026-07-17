using System;
using System.Collections.Generic;
using OfficeIMO.Latex;
using OfficeIMO.Latex.Markdown;
using OfficeIMO.Markdown.Pdf;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Latex.Pdf;

internal static class LatexPdfConversionEngine {
    internal static PdfCore.PdfDocumentConversionResult Convert(LatexDocument document, LatexPdfSaveOptions? options) {
        if (document == null) throw new ArgumentNullException(nameof(document));

        LatexPdfSaveOptions operation = (options ?? new LatexPdfSaveOptions()).CloneForConversion();
        LatexToMarkdownResult projection = document.ToMarkdownDocumentResult(operation.ProjectionOptions);
        PdfCore.PdfDocumentConversionResult result = projection.Value.ToPdfDocumentResult(operation.PdfOptions);
        return result.WithAdditionalWarnings(ToPdfWarnings(document, projection));
    }

    private static IEnumerable<PdfCore.PdfConversionWarning> ToPdfWarnings(
        LatexDocument document,
        LatexToMarkdownResult projection) {
        foreach (LatexDiagnostic diagnostic in document.Diagnostics) {
            yield return new PdfCore.PdfConversionWarning(
                "OfficeIMO.Latex.Pdf",
                diagnostic.Code,
                "parser @ " + diagnostic.Span,
                diagnostic.Message,
                ToPdfSeverity(diagnostic.Severity),
                details: new Dictionary<string, string> {
                    ["stage"] = "parse",
                    ["sourceSpan"] = diagnostic.Span.ToString()
                });
        }

        foreach (LatexMarkdownConversionDiagnostic diagnostic in projection.Report.Diagnostics) {
            string sourceSpan = diagnostic.LatexSpan.HasValue ? diagnostic.LatexSpan.Value.ToString() : "unknown";
            yield return new PdfCore.PdfConversionWarning(
                "OfficeIMO.Latex.Pdf",
                diagnostic.Code,
                diagnostic.Feature + " @ " + sourceSpan,
                diagnostic.Message,
                diagnostic.Outcome == LatexMarkdownConversionOutcome.Converted
                    ? PdfCore.PdfConversionWarningSeverity.Information
                    : PdfCore.PdfConversionWarningSeverity.Warning,
                details: new Dictionary<string, string> {
                    ["stage"] = "semantic-projection",
                    ["feature"] = diagnostic.Feature,
                    ["outcome"] = diagnostic.Outcome.ToString(),
                    ["sourceSpan"] = sourceSpan
                });
        }
    }

    private static PdfCore.PdfConversionWarningSeverity ToPdfSeverity(LatexDiagnosticSeverity severity) => severity switch {
        LatexDiagnosticSeverity.Error => PdfCore.PdfConversionWarningSeverity.Error,
        LatexDiagnosticSeverity.Warning => PdfCore.PdfConversionWarningSeverity.Warning,
        _ => PdfCore.PdfConversionWarningSeverity.Information
    };
}
