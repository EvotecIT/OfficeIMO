using OfficeIMO.AsciiDoc;
using OfficeIMO.AsciiDoc.Markdown;
using OfficeIMO.AsciiDoc.Pdf;
using OfficeIMO.Latex;
using OfficeIMO.Latex.Markdown;
using OfficeIMO.Latex.Pdf;
using OfficeIMO.Markdown;

namespace OfficeIMO.Workflows;

public static partial class OfficeWorkflow {
    private sealed partial class PairRoute {
        private async Task<OfficeConversionReport> ConvertTextAsync(Stream input, Stream output, OfficeConversionOptions options,
            OfficeWorkflowRouteSettings settings, CancellationToken token) {
            if (Id.StartsWith("asciidoc-", StringComparison.Ordinal)) {
                var read = AsciiDocDocument.LoadResult(input, settings.AsciiDocLoad, null, token);
                var source = new OfficeConversionReport("AsciiDoc", "AsciiDoc", diagnostics: read.Document.Diagnostics.Select(diagnostic =>
                    new OfficeDiagnostic(diagnostic.Code, Severity((int)diagnostic.Severity), diagnostic.Message, diagnostic.Span.ToString(),
                        "AsciiDoc", LossKind((int)diagnostic.Severity))));
                if (Id == "asciidoc-pdf") {
                    AsciiDocPdfExportOptions conversion = settings.AsciiDocPdf ?? new();
                    conversion.FailOnLoss |= options.FailOnLoss;
                    return await RetainSourceAsync(source, () => read.Document.SaveAsPdfAsync(output, conversion, token), conversion.FailOnLoss).ConfigureAwait(false);
                }
                AsciiDocMarkdownExportOptions projection = settings.AsciiDocMarkdown ?? new();
                projection.FailOnLoss |= options.FailOnLoss;
                return await RetainSourceAsync(source, () => read.Document.SaveAsMarkdownAsync(output, projection, token), projection.FailOnLoss).ConfigureAwait(false);
            }
            if (Id.StartsWith("latex-", StringComparison.Ordinal)) {
                var read = LatexDocument.LoadResult(input, settings.LatexLoad, null, token);
                var source = new OfficeConversionReport("LaTeX", "LaTeX", diagnostics: read.Document.Diagnostics.Select(diagnostic =>
                    new OfficeDiagnostic(diagnostic.Code, Severity((int)diagnostic.Severity), diagnostic.Message, diagnostic.Span.ToString(),
                        "LaTeX", LossKind((int)diagnostic.Severity))));
                if (Id == "latex-pdf") {
                    LatexPdfExportOptions conversion = settings.LatexPdf ?? new();
                    conversion.FailOnLoss |= options.FailOnLoss;
                    return await RetainSourceAsync(source, () => read.Document.SaveAsPdfAsync(output, conversion, token), conversion.FailOnLoss).ConfigureAwait(false);
                }
                LatexMarkdownExportOptions projection = settings.LatexMarkdown ?? new();
                projection.FailOnLoss |= options.FailOnLoss;
                return await RetainSourceAsync(source, () => read.Document.SaveAsMarkdownAsync(output, projection, token), projection.FailOnLoss).ConfigureAwait(false);
            }
            // AsciiDoc and LaTeX diagnostics locate Markdown losses by their source spans.
            using MarkdownDocument document = await LoadMarkdownAsync(input, options, settings, token, captureSyntaxTree: true).ConfigureAwait(false);
            if (Id == "markdown-asciidoc") {
                MarkdownAsciiDocExportOptions conversion = settings.MarkdownAsciiDoc ?? new();
                conversion.FailOnLoss |= options.FailOnLoss;
                return await document.SaveAsAsciiDocAsync(output, conversion, token).ConfigureAwait(false);
            }
            MarkdownLatexExportOptions latex = settings.MarkdownLatex ?? new();
            latex.FailOnLoss |= options.FailOnLoss;
            return await document.SaveAsLatexAsync(output, latex, token).ConfigureAwait(false);
        }

        private static Task<MarkdownDocument> LoadMarkdownAsync(Stream input, OfficeConversionOptions options,
            OfficeWorkflowRouteSettings settings, CancellationToken token, bool captureSyntaxTree = false) {
            // The document is converted and disposed; reader options the caller passed keep their own setting.
            MarkdownLoadOptions load = settings.MarkdownLoad ?? new() { CaptureSyntaxTree = captureSyntaxTree };
            load.MaxInputBytes = Math.Min(load.MaxInputBytes ?? options.MaximumInputBytes, options.MaximumInputBytes);
            return MarkdownDocument.LoadAsync(input, load, token);
        }

        private static OfficeDiagnosticSeverity Severity(int native) => native switch {
            2 => OfficeDiagnosticSeverity.Error, 1 => OfficeDiagnosticSeverity.Warning, _ => OfficeDiagnosticSeverity.Information
        };

        private static OfficeConversionLossKind LossKind(int native) => native switch {
            2 => OfficeConversionLossKind.Failure, 1 => OfficeConversionLossKind.Approximation, _ => OfficeConversionLossKind.None
        };
    }
}
