using OfficeIMO.Notebook;
using OfficeIMO.Notebook.Html;
using OfficeIMO.Notebook.Markdown;
using OfficeIMO.Notebook.Pdf;
using OfficeIMO.Notebook.Word;

namespace OfficeIMO;

public static partial class OfficeConverter {
    private static async Task<OfficeConversionReport> ConvertNotebookAsync(string id, Stream input,
        Stream output, OfficeConversionOptions settings, CancellationToken token) {
        var load = NotebookLoadOptions.UntrustedDefaults;
        load.MaxInputBytes = Math.Min(load.MaxInputBytes, settings.MaximumInputBytes);
        using NotebookDocument notebook = await NotebookDocument.LoadAsync(input, load, token).ConfigureAwait(false);
        switch (id) {
            case "notebook-pdf": {
                NotebookPdfExportOptions options = settings.NotebookPdf ?? new();
                options.FailOnLoss |= settings.FailOnLoss;
                return await notebook.SaveAsPdfAsync(output, options, token).ConfigureAwait(false);
            }
            case "notebook-html": {
                NotebookHtmlExportOptions options = settings.NotebookHtml ?? new();
                options.FailOnLoss |= settings.FailOnLoss;
                return await notebook.SaveAsHtmlAsync(output, options, token).ConfigureAwait(false);
            }
            case "notebook-docx": {
                NotebookWordExportOptions options = settings.NotebookWord ?? new();
                options.FailOnLoss |= settings.FailOnLoss;
                var converted = await notebook.ToWordAsync(options, token).ConfigureAwait(false);
                return await SaveConvertedAsync(converted,
                    word => word.SaveAsync(output, settings.WordSave, token)).ConfigureAwait(false);
            }
            case "notebook-markdown": {
                NotebookMarkdownExportOptions options = settings.NotebookMarkdown ?? new();
                options.FailOnLoss |= settings.FailOnLoss;
                using var markdown = notebook.ToMarkdown(options, out OfficeConversionReport report, token);
                await markdown.SaveAsync(output, cancellationToken: token).ConfigureAwait(false);
                return report;
            }
            default:
                throw new ArgumentOutOfRangeException(nameof(id), id, "The notebook conversion route is unknown.");
        }
    }
}
