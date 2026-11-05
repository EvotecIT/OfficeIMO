using OfficeIMO.Html;
using OfficeIMO.OneNote;
using OfficeIMO.OneNote.Html;
using OfficeIMO.OneNote.Markdown;
using OfficeIMO.OneNote.Pdf;

namespace OfficeIMO.Workflows;

public static partial class OfficeWorkflow {
    private sealed partial class PairRoute {
        private async Task<OfficeConversionReport> ConvertOneNoteAsync(Stream input, Stream output, OfficeConversionOptions options,
            OfficeWorkflowRouteSettings settings, CancellationToken token) {
            if (Id == "html-onenote") {
                HtmlConversionDocument document = await HtmlConversionDocument.LoadAsync(input, options.HtmlLoad, cancellationToken: token).ConfigureAwait(false);
                HtmlOneNoteExportOptions conversion = settings.HtmlOneNote ?? new();
                conversion.FailOnLoss |= options.FailOnLoss;
                return await SaveModelAsync(await document.ToOneNoteSectionAsync(conversion, token).ConfigureAwait(false), model => {
                    token.ThrowIfCancellationRequested();
                    model.Save(output, settings.OneNoteSave);
                    token.ThrowIfCancellationRequested();
                    return Task.CompletedTask;
                }, token).ConfigureAwait(false);
            }
            token.ThrowIfCancellationRequested();
            if (Path.GetExtension(options.SourceName ?? string.Empty).Equals(".onetoc2", StringComparison.OrdinalIgnoreCase)) {
                if (options.ResourceSourcePath == null)
                    throw new ArgumentException("Use a local .onetoc2 path or a self-contained .onepkg to convert notebook section content.");
                OneNoteNotebookReaderOptions reading = NotebookReading(options, settings);
                using OneNoteNotebookCapture captured = await OneNoteNotebookCapture.CreateAsync(input, options.ResourceSourcePath,
                    reading, options.MaximumInputBytes, token).ConfigureAwait(false);
                options.PublicationGuard = OfficePublicationGuards.Combine(options.PublicationGuard, captured.Guard);
                OneNoteNotebook notebook = captured.Read();
                token.ThrowIfCancellationRequested();
                OfficeConversionReport report = await ProjectOneNoteAsync(notebook, output, options, settings, token).ConfigureAwait(false);
                await captured.Guard.VerifyAsync(token).ConfigureAwait(false);
                return report;
            }
            if (Path.GetExtension(options.SourceName ?? string.Empty).Equals(".onepkg", StringComparison.OrdinalIgnoreCase)) {
                OneNoteNotebookReaderOptions reading = NotebookReading(options, settings);
                OneNoteNotebook notebook = OneNotePackageReader.Read(input, options.SourceName!, reading);
                token.ThrowIfCancellationRequested();
                return await ProjectOneNoteAsync(notebook, output, options, settings, token).ConfigureAwait(false);
            }
            OneNoteLoadOptions load = settings.OneNoteLoad ?? new();
            load.MaxInputBytes = Math.Min(load.MaxInputBytes ?? options.MaximumInputBytes, options.MaximumInputBytes);
            using OneNoteSection section = OneNoteSection.Load(input, load);
            token.ThrowIfCancellationRequested();
            return await ProjectOneNoteAsync(section, output, options, settings, token).ConfigureAwait(false);
        }

        private static OneNoteNotebookReaderOptions NotebookReading(OfficeConversionOptions options, OfficeWorkflowRouteSettings settings) {
            OneNoteNotebookReaderOptions reading = settings.OneNoteNotebookLoad ?? new();
            reading.OneNoteOptions = settings.OneNoteLoad ?? reading.OneNoteOptions;
            ArgumentNullException.ThrowIfNull(reading.OneNoteOptions);
            reading.OneNoteOptions.MaxInputBytes = Math.Min(reading.OneNoteOptions.MaxInputBytes ?? options.MaximumInputBytes, options.MaximumInputBytes);
            return reading;
        }

        private async Task<OfficeConversionReport> ProjectOneNoteAsync(object source, Stream output, OfficeConversionOptions options,
            OfficeWorkflowRouteSettings settings, CancellationToken token) {
            if (Id == "onenote-markdown") {
                OneNoteMarkdownExportOptions conversion = settings.OneNoteMarkdown ?? new();
                conversion.FailOnLoss |= options.FailOnLoss;
                return await (source is OneNoteSection markdownSection
                    ? markdownSection.SaveAsMarkdownAsync(output, conversion, token)
                    : ((OneNoteNotebook)source).SaveAsMarkdownAsync(output, conversion, token)).ConfigureAwait(false);
            }
            if (Id == "onenote-html") {
                if (settings.OneNoteVisualHtml != null) {
                    if (settings.OneNoteHtml != null || settings.OneNoteMarkdown != null || settings.OneNoteHtmlRendering != null)
                        throw new ArgumentException("Choose semantic or visual OneNote HTML options.");
                    settings.OneNoteVisualHtml.FailOnLoss |= options.FailOnLoss;
                    return await (source is OneNoteSection visualHtmlSection
                        ? visualHtmlSection.SaveAsVisualHtmlAsync(output, settings.OneNoteVisualHtml, token)
                        : ((OneNoteNotebook)source).SaveAsVisualHtmlAsync(output, settings.OneNoteVisualHtml, token)).ConfigureAwait(false);
                }
                OneNoteHtmlExportOptions conversion = settings.OneNoteHtml ?? new();
                conversion.FailOnLoss |= options.FailOnLoss;
                if (settings.OneNoteMarkdown != null) conversion.ProjectionOptions = settings.OneNoteMarkdown;
                if (settings.OneNoteHtmlRendering != null) conversion.HtmlOptions = settings.OneNoteHtmlRendering;
                return await (source is OneNoteSection htmlSection
                    ? htmlSection.SaveAsHtmlAsync(output, conversion, token)
                    : ((OneNoteNotebook)source).SaveAsHtmlAsync(output, conversion, token)).ConfigureAwait(false);
            }
            if (settings.OneNoteVisualPdf != null) {
                if (settings.OneNotePdf != null || settings.OneNoteMarkdown != null) throw new ArgumentException("Choose semantic or visual OneNote PDF options.");
                settings.OneNoteVisualPdf.FailOnLoss |= options.FailOnLoss;
                return await (source is OneNoteSection visualPdfSection
                    ? visualPdfSection.SaveAsVisualPdfAsync(output, settings.OneNoteVisualPdf, token)
                    : ((OneNoteNotebook)source).SaveAsVisualPdfAsync(output, settings.OneNoteVisualPdf, token)).ConfigureAwait(false);
            }
            OneNotePdfExportOptions pdf = settings.OneNotePdf ?? new();
            pdf.FailOnLoss |= options.FailOnLoss;
            if (settings.OneNoteMarkdown != null) {
                if (settings.OneNotePdf != null) throw new ArgumentException("Choose the OneNote PDF projection in Pdf options or separate projection options.");
                pdf.ProjectionOptions = settings.OneNoteMarkdown;
            }
            return await (source is OneNoteSection pdfSection
                ? pdfSection.SaveAsPdfAsync(output, pdf, token)
                : ((OneNoteNotebook)source).SaveAsPdfAsync(output, pdf, token)).ConfigureAwait(false);
        }
    }
}
