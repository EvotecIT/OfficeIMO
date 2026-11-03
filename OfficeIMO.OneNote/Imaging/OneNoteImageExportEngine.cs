using OfficeIMO.Drawing;
using System.Threading;

namespace OfficeIMO.OneNote;

internal static class OneNoteImageExportEngine {
    internal static IReadOnlyList<OfficeImageExportResult> ExportSection(
        OneNoteSection section,
        OfficeImageExportFormat format,
        OneNotePageBatchRenderingOptions options) {
        ArgumentNullException.ThrowIfNull(section);
        var results = new List<OfficeImageExportResult>();
        ExportSection(section, format, options, results.Add);
        return results.AsReadOnly();
    }

    internal static void ExportSection(
        OneNoteSection section,
        OfficeImageExportFormat format,
        OneNotePageBatchRenderingOptions options,
        OfficeImageExportConsumer consumer,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(section);
        ArgumentNullException.ThrowIfNull(consumer);
        OneNotePageReference[] pages = Select(OneNotePageTraversal.Flatten(section), options).ToArray();
        var encodingBudget = new OfficeImageExportEncodingBudget(options.MaximumTotalEncodedBytes);
        OfficeImageExportBatchProcessor.ForEachOrdered(
            pages,
            options.MaximumDegreeOfParallelism,
            (item, _, token) => OneNotePageImageRenderer.Render(
                item.Page,
                format,
                options,
                item.Page.Title,
                item.SectionPath + "/page[" + item.Index + "]",
                token,
                encodingBudget),
            consumer,
            cancellationToken,
            options);
    }

    internal static IReadOnlyList<OfficeImageExportResult> ExportNotebook(
        OneNoteNotebook notebook,
        OfficeImageExportFormat format,
        OneNotePageBatchRenderingOptions options) {
        ArgumentNullException.ThrowIfNull(notebook);
        var results = new List<OfficeImageExportResult>();
        ExportNotebook(notebook, format, options, results.Add);
        return results.AsReadOnly();
    }

    internal static void ExportNotebook(
        OneNoteNotebook notebook,
        OfficeImageExportFormat format,
        OneNotePageBatchRenderingOptions options,
        OfficeImageExportConsumer consumer,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(notebook);
        ArgumentNullException.ThrowIfNull(consumer);
        OneNotePageReference[] pages = Select(OneNotePageTraversal.Flatten(notebook), options).ToArray();
        var encodingBudget = new OfficeImageExportEncodingBudget(options.MaximumTotalEncodedBytes);
        OfficeImageExportBatchProcessor.ForEachOrdered(
            pages,
            options.MaximumDegreeOfParallelism,
            (item, _, token) => OneNotePageImageRenderer.Render(
                item.Page,
                format,
                options,
                item.Page.Title,
                notebook.Name + "/" + item.SectionPath + "/page[" + item.Index + "]",
                token,
                encodingBudget),
            consumer,
            cancellationToken,
            options);
    }

    private static IEnumerable<OneNotePageReference> Select(
        IReadOnlyList<OneNotePageReference> pages,
        OneNotePageBatchRenderingOptions options) {
        options.ValidateBatch();
        int end = options.PageCount.HasValue
            ? (int)Math.Min((long)pages.Count, (long)options.PageIndex + options.PageCount.Value)
            : pages.Count;
        for (int index = Math.Min(options.PageIndex, pages.Count); index < end; index++) {
            yield return pages[index];
        }
    }
}
