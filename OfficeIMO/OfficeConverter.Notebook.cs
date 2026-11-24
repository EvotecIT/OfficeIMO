using System.Text.Json;
using OfficeIMO.Notebook;
using OfficeIMO.Notebook.Html;
using OfficeIMO.Notebook.Import;
using OfficeIMO.Notebook.Markdown;
using OfficeIMO.Notebook.Pdf;
using OfficeIMO.Notebook.Word;

namespace OfficeIMO;

public static partial class OfficeConverter {
    private static async Task<OfficeConversionReport> ConvertNotebookAsync(string id, Stream input,
        Stream output, OfficeConversionOptions settings, CancellationToken token) {
        var load = NotebookLoadOptions.UntrustedDefaults;
        load.MaxInputBytes = Math.Min(load.MaxInputBytes, settings.MaximumInputBytes);
        if (id is "notebook-source-ipynb" or "notebook-source-pdf") {
            if (!NotebookImporter.TryGetFormatFromPath(settings.SourceName ?? string.Empty, out NotebookImportFormat format))
                throw new NotSupportedException("The notebook source extension has no implicit import profile.");
            var imported = await NotebookImporter.LoadAsync(input, format, new NotebookImportOptions { Limits = load }, token).ConfigureAwait(false);
            using NotebookDocument source = imported.Value;
            if (id == "notebook-source-pdf") {
                string? title = source.Metadata.TryGetValue("title", out JsonElement value) && value.ValueKind == JsonValueKind.String
                    ? value.GetString() : null;
                return await SaveNotebookPdfAsync(source, output, settings, token, title).ConfigureAwait(false);
            }
            if (settings.FailOnLoss) imported.ThrowIfLoss();
            await source.SaveAsync(output, new NotebookSaveOptions { AllowLoss = true }, token).ConfigureAwait(false);
            return imported.Report;
        }

        using NotebookDocument notebook = await NotebookDocument.LoadAsync(input, load, token).ConfigureAwait(false);
        switch (id) {
            case "notebook-pdf":
                return await SaveNotebookPdfAsync(notebook, output, settings, token).ConfigureAwait(false);
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
                return await notebook.SaveAsMarkdownAsync(output, options, token).ConfigureAwait(false);
            }
            default:
                throw new ArgumentOutOfRangeException(nameof(id), id, "The notebook conversion route is unknown.");
        }
    }

    private static Task<OfficeConversionReport> SaveNotebookPdfAsync(NotebookDocument notebook, Stream output,
        OfficeConversionOptions settings, CancellationToken token, string? sourceTitle = null) {
        NotebookPdfExportOptions options = settings.NotebookPdf ?? new();
        if (options.Title == null && options.MarkdownOptions.Title == null) options.Title = sourceTitle;
        options.FailOnLoss |= settings.FailOnLoss;
        return notebook.SaveAsPdfAsync(output, options, token);
    }
}
