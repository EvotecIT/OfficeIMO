using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

public static partial class HtmlImageExportExtensions {
    /// <summary>Renders one selected surface to the requested image format with dimensions and diagnostics.</summary>
    public static OfficeImageExportResult ExportImage(this HtmlConversionDocument document, OfficeImageExportFormat format, HtmlRenderOptions? options = null, int pageIndex = 0) {
        HtmlRenderOptions resolved = Normalize(options, pageIndex);
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(document, resolved);
        if (pageIndex >= rendered.Pages.Count) throw new ArgumentOutOfRangeException(nameof(pageIndex), "The selected HTML render page does not exist.");
        return RenderPage(rendered.Pages[pageIndex], format, resolved, rendered.DiagnosticReport, CancellationToken.None);
    }

    /// <summary>Renders all surfaces to the requested image format.</summary>
    public static IReadOnlyList<OfficeImageExportResult> ExportImages(this HtmlConversionDocument document, OfficeImageExportFormat format, HtmlRenderOptions? options = null) {
        var results = new List<OfficeImageExportResult>();
        document.ExportImages(format, results.Add, options);
        return results.AsReadOnly();
    }

    /// <summary>Streams rendered surfaces without retaining earlier encoded payloads.</summary>
    public static void ExportImages(
        this HtmlConversionDocument document,
        OfficeImageExportFormat format,
        OfficeImageExportConsumer consumer,
        HtmlRenderOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (consumer == null) throw new ArgumentNullException(nameof(consumer));
        cancellationToken.ThrowIfCancellationRequested();
        HtmlRenderOptions resolved = Normalize(options, 0);
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(document, resolved);
        OfficeImageExportBatchProcessor.ForEachOrdered(
            rendered.Pages,
            resolved.MaximumDegreeOfParallelism,
            (page, _, token) => RenderPage(page, format, resolved, rendered.DiagnosticReport, token),
            consumer,
            cancellationToken,
            resolved);
    }

    /// <summary>Asynchronously renders one selected surface to the requested image format.</summary>
    public static async Task<OfficeImageExportResult> ExportImageAsync(this HtmlConversionDocument document, OfficeImageExportFormat format, HtmlRenderOptions? options = null, int pageIndex = 0, CancellationToken cancellationToken = default) {
        HtmlRenderOptions resolved = Normalize(options, pageIndex);
        HtmlRenderDocument rendered = await HtmlRenderEngine.RenderAsync(document, resolved, cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        if (pageIndex >= rendered.Pages.Count) throw new ArgumentOutOfRangeException(nameof(pageIndex), "The selected HTML render page does not exist.");
        return RenderPage(rendered.Pages[pageIndex], format, resolved, rendered.DiagnosticReport, cancellationToken);
    }

    /// <summary>Asynchronously renders all surfaces to the requested image format.</summary>
    public static async Task<IReadOnlyList<OfficeImageExportResult>> ExportImagesAsync(this HtmlConversionDocument document, OfficeImageExportFormat format, HtmlRenderOptions? options = null, CancellationToken cancellationToken = default) {
        var results = new List<OfficeImageExportResult>();
        await document.ExportImagesAsync(
            format,
            (result, _) => {
                results.Add(result);
                return Task.CompletedTask;
            },
            options,
            cancellationToken).ConfigureAwait(false);
        return results.AsReadOnly();
    }

    /// <summary>Asynchronously streams rendered surfaces without retaining earlier encoded payloads.</summary>
    public static async Task ExportImagesAsync(
        this HtmlConversionDocument document,
        OfficeImageExportFormat format,
        OfficeImageExportAsyncConsumer consumer,
        HtmlRenderOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (consumer == null) throw new ArgumentNullException(nameof(consumer));
        HtmlRenderOptions resolved = Normalize(options, 0);
        HtmlRenderDocument rendered = await HtmlRenderEngine.RenderAsync(document, resolved, cancellationToken).ConfigureAwait(false);
        OfficeImageExportAsyncConsumer accept =
            OfficeImageExportBatchProcessor.CreateGuardedAsyncConsumer(
                resolved,
                consumer,
                cancellationToken);
        foreach (HtmlRenderPage page in rendered.Pages) {
            cancellationToken.ThrowIfCancellationRequested();
            OfficeImageExportResult result = RenderPage(page, format, resolved, rendered.DiagnosticReport, cancellationToken);
            await accept(result, cancellationToken).ConfigureAwait(false);
        }
    }

}
