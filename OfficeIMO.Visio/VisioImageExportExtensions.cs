using OfficeIMO.Drawing;

namespace OfficeIMO.Visio;

/// <summary>Format-neutral image export for Visio documents and pages.</summary>
public static partial class VisioImageExportExtensions {
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
        ArgumentNullException.ThrowIfNull(document);
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
        ArgumentNullException.ThrowIfNull(document);
        ArgumentNullException.ThrowIfNull(consumer);
        VisioImageExportOptions resolved = Normalize(options);
        if (document.Pages.Count == 0) throw new InvalidOperationException("The document does not contain any pages to export.");
        if (resolved.PageIndex >= document.Pages.Count) {
            throw new ArgumentOutOfRangeException(nameof(options), "PageIndex is outside the document page collection.");
        }

        int available = document.Pages.Count - resolved.PageIndex;
        int count = resolved.PageCount.HasValue ? Math.Min(resolved.PageCount.Value, available) : available;
        int[] pageIndexes = Enumerable.Range(resolved.PageIndex, count).ToArray();
        var encodingBudget = new OfficeImageExportEncodingBudget(resolved.MaximumTotalEncodedBytes);
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
                    token,
                    encodingBudget);
            },
            consumer,
            cancellationToken,
            resolved);
    }

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
