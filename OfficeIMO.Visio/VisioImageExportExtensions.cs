using OfficeIMO.Drawing;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Visio;

/// <summary>Format-neutral image export for Visio documents and pages.</summary>
public static class VisioImageExportExtensions {
    /// <summary>Exports a Visio page to a supported raster format or SVG.</summary>
    public static OfficeImageExportResult ExportImage(
        this VisioPage page,
        OfficeImageExportFormat format,
        VisioImageExportOptions? options = null,
        CancellationToken cancellationToken = default) =>
        VisioImageExportEngine.Render(page, format, Normalize(options), cancellationToken: cancellationToken);

    /// <summary>Exports the selected document page to a supported raster format or SVG.</summary>
    public static OfficeImageExportResult ExportImage(
        this VisioDocument document,
        OfficeImageExportFormat format,
        VisioImageExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        VisioImageExportOptions resolved = Normalize(options);
        if (document.Pages.Count == 0) throw new InvalidOperationException("The document does not contain any pages to export.");
        if (resolved.PageIndex >= document.Pages.Count) {
            throw new ArgumentOutOfRangeException(nameof(options), "PageIndex is outside the document page collection.");
        }
        int pageNumber = resolved.PageIndex + 1;
        VisioPage page = document.Pages[resolved.PageIndex];
        return VisioImageExportEngine.Render(
            page,
            format,
            resolved,
            ResolvePageName(page, pageNumber),
            "Visio page " + pageNumber,
            cancellationToken);
    }

    /// <summary>Exports a selected range of document pages to a supported raster format or SVG.</summary>
    public static IReadOnlyList<OfficeImageExportResult> ExportImages(
        this VisioDocument document,
        OfficeImageExportFormat format,
        VisioImageExportOptions? options = null) {
        var results = new List<OfficeImageExportResult>();
        document.ExportImages(format, results.Add, options);
        return results.AsReadOnly();
    }

    /// <summary>Streams selected document page images without retaining earlier payloads.</summary>
    public static void ExportImages(
        this VisioDocument document,
        OfficeImageExportFormat format,
        OfficeImageExportConsumer consumer,
        VisioImageExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (consumer == null) throw new ArgumentNullException(nameof(consumer));
        VisioImageExportOptions resolved = Normalize(options);
        if (document.Pages.Count == 0) throw new InvalidOperationException("The document does not contain any pages to export.");
        if (resolved.PageIndex >= document.Pages.Count) {
            throw new ArgumentOutOfRangeException(nameof(options), "PageIndex is outside the document page collection.");
        }

        int available = document.Pages.Count - resolved.PageIndex;
        int count = resolved.PageCount.HasValue ? Math.Min(resolved.PageCount.Value, available) : available;
        int[] pageIndexes = Enumerable.Range(resolved.PageIndex, count).ToArray();
        OfficeImageExportBatchProcessor.ForEachOrdered(
            pageIndexes,
            resolved.MaximumDegreeOfParallelism,
            (pageIndex, _, token) => {
                int pageNumber = pageIndex + 1;
                VisioPage page = document.Pages[pageIndex];
                return VisioImageExportEngine.Render(
                    page,
                    format,
                    resolved,
                    ResolvePageName(page, pageNumber),
                    "Visio page " + pageNumber,
                    token);
            },
            consumer,
            cancellationToken,
            resolved);
    }

    /// <summary>Renders a Visio page to JPEG bytes.</summary>
    public static byte[] ToJpeg(this VisioPage page, VisioImageExportOptions? options = null) =>
        page.ExportImage(OfficeImageExportFormat.Jpeg, options).Bytes;

    /// <summary>Renders the selected Visio document page to JPEG bytes.</summary>
    public static byte[] ToJpeg(this VisioDocument document, VisioImageExportOptions? options = null) =>
        document.ExportImage(OfficeImageExportFormat.Jpeg, options).Bytes;

    /// <summary>Renders a Visio page to TIFF bytes.</summary>
    public static byte[] ToTiff(this VisioPage page, VisioImageExportOptions? options = null) =>
        page.ExportImage(OfficeImageExportFormat.Tiff, options).Bytes;

    /// <summary>Renders the selected Visio document page to TIFF bytes.</summary>
    public static byte[] ToTiff(this VisioDocument document, VisioImageExportOptions? options = null) =>
        document.ExportImage(OfficeImageExportFormat.Tiff, options).Bytes;

    /// <summary>Renders a Visio page to lossless WebP bytes.</summary>
    public static byte[] ToWebp(this VisioPage page, VisioImageExportOptions? options = null) =>
        page.ExportImage(OfficeImageExportFormat.Webp, options).Bytes;

    /// <summary>Renders the selected Visio document page to lossless WebP bytes.</summary>
    public static byte[] ToWebp(this VisioDocument document, VisioImageExportOptions? options = null) =>
        document.ExportImage(OfficeImageExportFormat.Webp, options).Bytes;

    /// <summary>Saves selected document pages to a folder as PNG files.</summary>
    public static IReadOnlyList<OfficeImageExportResult> SaveAsImages(
        this VisioDocument document,
        string folderPath,
        VisioImageExportOptions? options = null) =>
        new VisioDocumentImageBatchExportBuilder(document, options).AsPng().Save(folderPath);

    /// <summary>Saves selected document pages to a folder in the requested image format.</summary>
    public static IReadOnlyList<OfficeImageExportResult> SaveAsImages(
        this VisioDocument document,
        string folderPath,
        OfficeImageExportFormat format,
        VisioImageExportOptions? options = null) =>
        new VisioDocumentImageBatchExportBuilder(document, options).As(format).Save(folderPath);

    /// <summary>Asynchronously saves selected document pages to a folder as PNG files.</summary>
    public static Task<IReadOnlyList<OfficeImageExportResult>> SaveAsImagesAsync(
        this VisioDocument document,
        string folderPath,
        VisioImageExportOptions? options = null,
        CancellationToken cancellationToken = default) =>
        new VisioDocumentImageBatchExportBuilder(document, options).AsPng().SaveAsync(folderPath, cancellationToken);

    /// <summary>Asynchronously saves selected document pages to a folder in the requested image format.</summary>
    public static Task<IReadOnlyList<OfficeImageExportResult>> SaveAsImagesAsync(
        this VisioDocument document,
        string folderPath,
        OfficeImageExportFormat format,
        VisioImageExportOptions? options = null,
        CancellationToken cancellationToken = default) =>
        new VisioDocumentImageBatchExportBuilder(document, options).As(format).SaveAsync(folderPath, cancellationToken);

    private static VisioImageExportOptions Normalize(VisioImageExportOptions? options) {
        VisioImageExportOptions resolved = options?.Clone() ?? new VisioImageExportOptions();
        resolved.Validate();
        return resolved;
    }

    private static string ResolvePageName(VisioPage page, int pageNumber) =>
        string.IsNullOrWhiteSpace(page.Name)
            ? "Page " + pageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture)
            : page.Name;
}
