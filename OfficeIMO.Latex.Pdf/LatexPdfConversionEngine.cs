using System;
using System.Collections.Generic;
using OfficeIMO.Latex;
using OfficeIMO.Latex.Markdown;
using OfficeIMO.Markdown.Pdf;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Latex.Pdf;

internal static class LatexPdfConversionEngine {
    internal static PdfCore.PdfDocumentConversionResult Convert(LatexDocument document, LatexToPdfOptions? options) {
        if (document == null) throw new ArgumentNullException(nameof(document));

        LatexToPdfOptions operation = (options ?? new LatexToPdfOptions()).CloneForConversion();
        LatexToMarkdownResult projection = document.ToMarkdownDocumentResult(operation.ProjectionOptions);
        PdfCore.PdfDocumentConversionResult result = projection.Value.ToPdfDocumentResult(operation.MarkdownOptions);
        return result
            .WithSourceConversionReport(projection.Report)
            .WithAdditionalWarnings(ToPdfWarnings(document));
    }

    private static IEnumerable<PdfCore.PdfConversionWarning> ToPdfWarnings(LatexDocument document) {
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

    }

    private static PdfCore.PdfConversionWarningSeverity ToPdfSeverity(LatexDiagnosticSeverity severity) => severity switch {
        LatexDiagnosticSeverity.Error => PdfCore.PdfConversionWarningSeverity.Error,
        LatexDiagnosticSeverity.Warning => PdfCore.PdfConversionWarningSeverity.Warning,
        _ => PdfCore.PdfConversionWarningSeverity.Information
    };
}
