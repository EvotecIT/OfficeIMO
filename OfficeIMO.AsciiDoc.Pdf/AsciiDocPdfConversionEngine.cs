using System;
using System.Collections.Generic;
using OfficeIMO.AsciiDoc;
using OfficeIMO.AsciiDoc.Markdown;
using OfficeIMO.Markdown.Pdf;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.AsciiDoc.Pdf;

internal static class AsciiDocPdfConversionEngine {
    internal static PdfCore.PdfDocumentConversionResult Convert(AsciiDocDocument document, AsciiDocToPdfOptions? options) {
        if (document == null) throw new ArgumentNullException(nameof(document));

        AsciiDocToPdfOptions operation = (options ?? new AsciiDocToPdfOptions()).CloneForConversion();
        AsciiDocToMarkdownResult projection = document.ToMarkdownDocumentResult(operation.ProjectionOptions);
        PdfCore.PdfDocumentConversionResult result = projection.Value.ToPdfDocumentResult(operation.MarkdownOptions);
        return result
            .WithSourceConversionReport(projection.Report)
            .WithAdditionalWarnings(ToPdfWarnings(document));
    }

    private static IEnumerable<PdfCore.PdfConversionWarning> ToPdfWarnings(AsciiDocDocument document) {
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

    }

    private static PdfCore.PdfConversionWarningSeverity ToPdfSeverity(AsciiDocDiagnosticSeverity severity) => severity switch {
        AsciiDocDiagnosticSeverity.Error => PdfCore.PdfConversionWarningSeverity.Error,
        AsciiDocDiagnosticSeverity.Warning => PdfCore.PdfConversionWarningSeverity.Warning,
        _ => PdfCore.PdfConversionWarningSeverity.Information
    };
}
