using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

public static partial class HtmlImageExportExtensions {
    /// <summary>Renders one selected surface to the requested image format with dimensions and diagnostics.</summary>
    public static OfficeImageExportResult ExportImage(this HtmlConversionDocument document, OfficeImageExportFormat format, HtmlRenderOptions? options = null, int pageIndex = 0) {
        HtmlRenderOptions resolved = Normalize(options, pageIndex);
        return HtmlRenderEngine.ExecuteWithDeadline(resolved, CancellationToken.None, operationCancellationToken => {
            HtmlRenderDocument rendered = HtmlRenderEngine.Render(document, resolved, operationCancellationToken);
            if (pageIndex >= rendered.Pages.Count) throw new ArgumentOutOfRangeException(nameof(pageIndex), "The selected HTML render page does not exist.");
            return RenderPage(rendered.Pages[pageIndex], format, resolved, rendered.DiagnosticReport, operationCancellationToken);
        });
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
        HtmlRenderEngine.ExecuteWithDeadline(resolved, cancellationToken, operationCancellationToken => {
            HtmlRenderDocument rendered = HtmlRenderEngine.Render(
                document,
                resolved,
                operationCancellationToken);
            OfficeImageExportBatchProcessor.ForEachOrdered(
                rendered.Pages,
                resolved.MaximumDegreeOfParallelism,
                (page, _, token) => RenderPage(page, format, resolved, rendered.DiagnosticReport, token),
                consumer,
                operationCancellationToken,
                resolved);
            return true;
        });
    }

    /// <summary>Asynchronously renders one selected surface to the requested image format.</summary>
    public static async Task<OfficeImageExportResult> ExportImageAsync(this HtmlConversionDocument document, OfficeImageExportFormat format, HtmlRenderOptions? options = null, int pageIndex = 0, CancellationToken cancellationToken = default) {
        HtmlRenderOptions resolved = Normalize(options, pageIndex);
        return await HtmlRenderEngine.ExecuteWithDeadlineAsync(resolved, cancellationToken, async operationCancellationToken => {
            HtmlRenderDocument rendered = await HtmlRenderEngine.RenderAsync(document, resolved, operationCancellationToken).ConfigureAwait(false);
            operationCancellationToken.ThrowIfCancellationRequested();
            if (pageIndex >= rendered.Pages.Count) throw new ArgumentOutOfRangeException(nameof(pageIndex), "The selected HTML render page does not exist.");
            return RenderPage(rendered.Pages[pageIndex], format, resolved, rendered.DiagnosticReport, operationCancellationToken);
        }).ConfigureAwait(false);
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
        HtmlRenderDocument? rendered = null;
        await OfficeImageExportBatchProcessor.RunAsyncWithPreflight(
            resolved,
            async operationCancellationToken => {
                rendered = await HtmlRenderEngine.RenderAsync(
                    document,
                    resolved,
                    operationCancellationToken).ConfigureAwait(false);
                operationCancellationToken.ThrowIfCancellationRequested();
                return rendered.Pages.Count;
            },
            async (accept, operationCancellationToken) => {
                HtmlRenderDocument completed = rendered!;
                foreach (HtmlRenderPage page in completed.Pages) {
                    operationCancellationToken.ThrowIfCancellationRequested();
                    OfficeImageExportResult result = RenderPage(
                        page,
                        format,
                        resolved,
                        completed.DiagnosticReport,
                        operationCancellationToken);
                    await accept(result, operationCancellationToken).ConfigureAwait(false);
                }
            },
            consumer,
            cancellationToken).ConfigureAwait(false);
    }

}
