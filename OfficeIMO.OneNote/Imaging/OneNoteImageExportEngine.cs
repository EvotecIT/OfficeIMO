using OfficeIMO.Drawing;
using System.Threading;

namespace OfficeIMO.OneNote;

internal static class OneNoteImageExportEngine {
    internal static IReadOnlyList<OfficeImageExportResult> ExportSection(
        OneNoteSection section,
        OfficeImageExportFormat format,
        OneNotePageBatchRenderingOptions options) {
        if (section == null) throw new ArgumentNullException(nameof(section));
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
        if (section == null) throw new ArgumentNullException(nameof(section));
        if (consumer == null) throw new ArgumentNullException(nameof(consumer));
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
        if (notebook == null) throw new ArgumentNullException(nameof(notebook));
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
        if (notebook == null) throw new ArgumentNullException(nameof(notebook));
        if (consumer == null) throw new ArgumentNullException(nameof(consumer));
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
