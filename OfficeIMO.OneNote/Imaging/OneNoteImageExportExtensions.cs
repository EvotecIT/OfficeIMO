using OfficeIMO.Drawing;
using System.Threading;

namespace OfficeIMO.OneNote;

/// <summary>OneNote page rendering and image-export entry points.</summary>
public static partial class OneNoteImageExportExtensions {
    /// <summary>Exports a page using the shared five-format result contract.</summary>
    public static OfficeImageExportResult ExportImage(
        this OneNotePage page,
        OfficeImageExportFormat format,
        OneNotePageRenderingOptions? options = null,
        CancellationToken cancellationToken = default) =>
        OneNotePageImageRenderer.Render(
            page,
            format,
            options?.Clone() ?? new OneNotePageRenderingOptions(),
            cancellationToken: cancellationToken);

    /// <summary>Exports selected section pages using the shared five-format result contract.</summary>
    public static IReadOnlyList<OfficeImageExportResult> ExportImages(
        this OneNoteSection section,
        OfficeImageExportFormat format,
        OneNotePageBatchRenderingOptions? options = null) =>
        OneNoteImageExportEngine.ExportSection(
            section,
            format,
            options?.CloneBatch() ?? new OneNotePageBatchRenderingOptions());

    /// <summary>Exports selected notebook pages using the shared five-format result contract.</summary>
    public static IReadOnlyList<OfficeImageExportResult> ExportImages(
        this OneNoteNotebook notebook,
        OfficeImageExportFormat format,
        OneNotePageBatchRenderingOptions? options = null) =>
        OneNoteImageExportEngine.ExportNotebook(
            notebook,
            format,
            options?.CloneBatch() ?? new OneNotePageBatchRenderingOptions());

    /// <summary>Creates the reusable Drawing scene for a page.</summary>
    public static OfficeDrawing ToDrawing(this OneNotePage page, OneNotePageRenderingOptions? options = null) => OneNotePageRenderer.Render(page, options);
}
