using System.Text;
using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows;

public static partial class OfficeWorkflow {
    private sealed partial class PairRoute {
        private static async Task<OfficeConversionReport> ConvertPdfMarkdownAsync(Stream input, Stream output,
            OfficeConversionOptions options, OfficeWorkflowRouteSettings settings, CancellationToken token) {
            PdfLoadOptions load = (options.PdfLoad ?? new()).Clone(options.MaximumInputBytes,
                options.PdfLoad?.Password ?? options.SourcePassword);
            using PdfDocument document = await PdfDocument.LoadAsync(input, load, token).ConfigureAwait(false);
            string markdown = document.ExtractMarkdown(settings.PdfMarkdown, token);
            token.ThrowIfCancellationRequested();
            var report = new OfficeConversionReport("PDF", "Markdown", diagnostics: new[] {
                new OfficeDiagnostic("PDF_MARKDOWN_LOGICAL_PROJECTION", OfficeDiagnosticSeverity.Warning,
                    "Markdown retains reconstructed logical content; fixed page layout and interactive PDF behavior are not retained.",
                    lossKind: OfficeConversionLossKind.Approximation)
            });
            if (options.FailOnLoss) report.ThrowIfLoss();
            using var writer = new StreamWriter(output, new UTF8Encoding(false), 4096, leaveOpen: true);
            await writer.WriteAsync(markdown).ConfigureAwait(false);
            await writer.FlushAsync().ConfigureAwait(false);
            token.ThrowIfCancellationRequested();
            return report;
        }
    }
}
