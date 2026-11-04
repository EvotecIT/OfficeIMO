using OfficeIMO.Core.Internal;
using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows;

public static partial class OfficeWorkflow {
    private static IEnumerable<string> PairRouteIds => new[] {
        "asciidoc-markdown", "markdown-asciidoc", "latex-markdown", "markdown-latex", "markdown-pptx",
        "rtf-html", "html-rtf", "rtf-markdown", "markdown-rtf", "rtf-docx", "docx-rtf", "pdf-rtf",
        "docx-odt", "odt-docx", "xlsx-ods", "ods-xlsx", "pptx-odp", "odp-pptx", "xlsx-html", "html-xlsx",
        "onenote-html", "html-onenote", "onenote-markdown", "onenote-pdf",
        "asciidoc-pdf", "latex-pdf", "odt-pdf", "ods-pdf", "odp-pdf"
    };

    // The optional bundle owns only load/convert/save orchestration. Format mapping and loss policy stay in each pair adapter.
    private sealed partial class PairRoute(string id) : IOfficeConversionRoute, IOfficeConversionArtifactValidation,
        IOfficeConversionOutputValidation {
        public string Id => id;

        public async Task<OfficeConversionReport> ConvertAsync(Stream source, Stream destination, OfficeConversionOptions options,
            CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            OfficeWorkflowRouteSettings settings = CaptureSettings(options);
            if (options.OutputProfile != OfficeConversionProfile.Faithful)
                throw new ArgumentException("The selected pair supports the Faithful output profile.", nameof(options));
            OfficeConversionReport report = Id switch {
                "markdown-pptx" => await ConvertMarkdownPowerPointAsync(source, destination, options, settings, cancellationToken).ConfigureAwait(false),
                "xlsx-html" or "html-xlsx" => await ConvertExcelHtmlAsync(source, destination, options, settings, cancellationToken).ConfigureAwait(false),
                "asciidoc-markdown" or "markdown-asciidoc" or "latex-markdown" or "markdown-latex" or "asciidoc-pdf" or "latex-pdf" =>
                    await ConvertTextAsync(source, destination, options, settings, cancellationToken).ConfigureAwait(false),
                "rtf-html" or "html-rtf" or "rtf-markdown" or "markdown-rtf" or "rtf-docx" or "docx-rtf" or "pdf-rtf" =>
                    await ConvertRtfAsync(source, destination, options, settings, cancellationToken).ConfigureAwait(false),
                "docx-odt" or "odt-docx" or "xlsx-ods" or "ods-xlsx" or "pptx-odp" or "odp-pptx" or "odt-pdf" or "ods-pdf" or "odp-pdf" =>
                    await ConvertOpenDocumentAsync(source, destination, options, settings, cancellationToken).ConfigureAwait(false),
                "onenote-html" or "html-onenote" or "onenote-markdown" or "onenote-pdf" =>
                    await ConvertOneNoteAsync(source, destination, options, settings, cancellationToken).ConfigureAwait(false),
                _ => throw new InvalidOperationException("The pair route has no implementation.")
            };
            if (options.FailOnLoss) report.ThrowIfLoss();
            return report;
        }

        private OfficeWorkflowRouteSettings CaptureSettings(OfficeConversionOptions options) => options.RegisteredSettings switch {
            null => new OfficeWorkflowRouteSettings(),
            OfficeWorkflowRouteSettings settings => settings.ForRoute(Id),
            _ => throw new ArgumentException("The selected pair requires OfficeWorkflowRouteSettings.", nameof(options))
        };

        private static async Task<OfficeConversionReport> ConvertMarkdownPowerPointAsync(Stream input, Stream output,
            OfficeConversionOptions options, OfficeWorkflowRouteSettings settings, CancellationToken token) {
            MarkdownToPowerPointOptions conversion = settings.MarkdownPowerPoint ?? new();
            conversion.FailOnLoss |= options.FailOnLoss;
            conversion.BaseDirectory ??= options.ResourceDirectory;
            var load = conversion.ReaderOptions?.Clone() ?? OfficeIMO.Markdown.MarkdownLoadOptions.CreateOfficeIMOProfile();
            load.MaxInputBytes = Math.Min(load.MaxInputBytes ?? options.MaximumInputBytes, options.MaximumInputBytes);
            using var document = await OfficeIMO.Markdown.MarkdownDocument.LoadAsync(input, load, token).ConfigureAwait(false);
            return await SaveModelAsync(await document.ToPowerPointAsync(conversion, token).ConfigureAwait(false),
                model => model.SaveAsync(output, token), token).ConfigureAwait(false);
        }

        private static async Task<OfficeConversionReport> SaveModelAsync<T>(OfficeConversionResult<T> converted,
            Func<T, Task> save, CancellationToken cancellationToken) {
            try {
                cancellationToken.ThrowIfCancellationRequested();
                await OfficeConversionFailures.RunAsync(converted.Report, () => save(converted.Value)).ConfigureAwait(false);
                return converted.Report;
            } finally { (converted.Value as IDisposable)?.Dispose(); }
        }

        private static async Task<OfficeConversionReport> RetainSourceAsync(OfficeConversionReport source,
            Func<Task<OfficeConversionReport>> convert, bool failOnLoss) {
            if (failOnLoss) source.ThrowIfLoss();
            try { return Combine(source, await convert().ConfigureAwait(false)); }
            catch (OfficeConversionException error) {
                throw new OfficeConversionException(error.Message, Combine(source, error.Report), error);
            }
        }

        private static OfficeConversionReport Combine(OfficeConversionReport source, OfficeConversionReport target) =>
            new OfficeConversionReport(source.Source, target.Target, target.OutputPath, target.ItemsWritten, target.Elapsed,
                OfficeConversionFidelityDiagnostics.Flatten(new[] { source, target }));

        public PdfCreateOptions? GetOutputCreationOptions(OfficeConversionOptions options) {
            if (options.RegisteredSettings is not OfficeWorkflowRouteSettings settings) return null;
            return Id switch {
                "asciidoc-pdf" => settings.AsciiDocPdf?.MarkdownOptions.PdfOptions,
                "latex-pdf" => settings.LatexPdf?.MarkdownOptions.PdfOptions,
                "odt-pdf" => settings.OdtPdf?.PdfOptions?.PdfOptions,
                "ods-pdf" => settings.OdsPdf?.PdfOptions?.PdfOptions,
                "odp-pdf" => settings.OdpPdf?.PdfOptions?.PdfOptions,
                "onenote-pdf" => settings.OneNoteVisualPdf?.PdfOptions ?? settings.OneNotePdf?.MarkdownOptions.PdfOptions,
                _ => null
            };
        }

        public PdfLoadOptions? CreateOutputPdfLoadOptions(OfficeConversionOptions options) {
            PdfCreateOptions? creation = GetOutputCreationOptions(options);
            if (creation?.Encryption == null) return null;
            var load = new PdfLoadOptions { Password = creation.Encryption.OwnerPassword ?? creation.Encryption.UserPassword };
            return creation.Encryption.AesCryptographyProvider is { } provider
                ? PdfLoadOptions.WithAesCryptographyProvider(load, provider) : load;
        }
    }
}
