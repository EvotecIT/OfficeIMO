using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

public static partial class HtmlImageExportExtensions {
    /// <summary>Renders one selected surface to the requested image format with dimensions and diagnostics.</summary>
    public static OfficeImageExportResult ExportImage(this HtmlConversionDocument document, OfficeImageExportFormat format, HtmlRenderOptions? options = null, int pageIndex = 0) {
        HtmlRenderRequest request = HtmlRenderRequest.FromLegacy(
            options, MapEncoder(format), HtmlRenderPageSet.Page(pageIndex));
        HtmlRenderOptions resolved = HtmlRenderEngine.PrepareOptions(document, request);
        return HtmlRenderEngine.ExecuteWithDeadline(resolved, CancellationToken.None, operationCancellationToken => {
            HtmlRenderResult rendered = HtmlRenderEngine.ExecuteCore(document, request, resolved, operationCancellationToken);
            OfficeImageExportResult? image = null;
            HtmlRenderResultImageExtensions.ExportImagesCore(
                rendered, format, resolved, result => image = result, operationCancellationToken);
            return image ?? throw new InvalidOperationException("The selected render surface did not produce an image.");
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
        HtmlRenderRequest request = HtmlRenderRequest.FromLegacy(
            options, MapEncoder(format), HtmlRenderPageSet.All());
        HtmlRenderOptions resolved = HtmlRenderEngine.PrepareOptions(document, request);
        HtmlRenderEngine.ExecuteWithDeadline(resolved, cancellationToken, operationCancellationToken => {
            HtmlRenderResult rendered = HtmlRenderEngine.ExecuteCore(document, request, resolved, operationCancellationToken);
            HtmlRenderResultImageExtensions.ExportImagesCore(
                rendered, format, resolved, consumer, operationCancellationToken);
            return true;
        });
    }

    /// <summary>Asynchronously renders one selected surface to the requested image format.</summary>
    public static async Task<OfficeImageExportResult> ExportImageAsync(this HtmlConversionDocument document, OfficeImageExportFormat format, HtmlRenderOptions? options = null, int pageIndex = 0, CancellationToken cancellationToken = default) {
        HtmlRenderRequest request = HtmlRenderRequest.FromLegacy(
            options, MapEncoder(format), HtmlRenderPageSet.Page(pageIndex));
        HtmlRenderOptions resolved = HtmlRenderEngine.PrepareOptions(document, request);
        return await HtmlRenderEngine.ExecuteWithDeadlineAsync(resolved, cancellationToken, async operationCancellationToken => {
            HtmlRenderResult rendered = await HtmlRenderEngine.ExecuteCoreAsync(
                document, request, resolved, operationCancellationToken).ConfigureAwait(false);
            OfficeImageExportResult? image = null;
            HtmlRenderResultImageExtensions.ExportImagesCore(
                rendered, format, resolved, result => image = result, operationCancellationToken);
            return image ?? throw new InvalidOperationException("The selected render surface did not produce an image.");
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
        HtmlRenderRequest request = HtmlRenderRequest.FromLegacy(
            options, MapEncoder(format), HtmlRenderPageSet.All());
        HtmlRenderOptions resolved = HtmlRenderEngine.PrepareOptions(document, request);
        var encodingBudget = new OfficeImageExportEncodingBudget(resolved.MaximumTotalEncodedBytes);
        HtmlRenderResult? rendered = null;
        await OfficeImageExportBatchProcessor.RunAsyncWithPreflight(
            resolved,
            async operationCancellationToken => {
                rendered = await HtmlRenderEngine.ExecuteCoreAsync(document, request, resolved, operationCancellationToken).ConfigureAwait(false);
                operationCancellationToken.ThrowIfCancellationRequested();
                return rendered.Document.Pages.Count;
            },
            async (accept, operationCancellationToken) => {
                HtmlRenderResult completed = rendered!;
                foreach (HtmlRenderPage page in completed.Document.Pages) {
                    operationCancellationToken.ThrowIfCancellationRequested();
                    OfficeImageExportResult result = RenderPage(
                        page,
                        format,
                        resolved,
                        completed.Document.DiagnosticReport,
                        operationCancellationToken,
                        encodingBudget);
                    await accept(result, operationCancellationToken).ConfigureAwait(false);
                }
            },
            consumer,
            cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Executes an explicit request and encodes every retained image surface.</summary>
    public static IReadOnlyList<OfficeImageExportResult> RenderImages(
        this HtmlConversionDocument document,
        HtmlRenderRequest request,
        CancellationToken cancellationToken = default) {
        HtmlRenderOptions resolved = HtmlRenderEngine.PrepareOptions(document, request);
        return HtmlRenderEngine.ExecuteWithDeadline(resolved, cancellationToken, operationCancellationToken => {
            HtmlRenderResult rendered = HtmlRenderEngine.ExecuteCore(document, request, resolved, operationCancellationToken);
            var images = new List<OfficeImageExportResult>(rendered.Document.Pages.Count);
            HtmlRenderResultImageExtensions.ExportImagesCore(
                rendered, MapEncoderFormat(request.Encoder), resolved, images.Add, operationCancellationToken);
            return (IReadOnlyList<OfficeImageExportResult>)images.AsReadOnly();
        });
    }

    /// <summary>Executes an explicit request asynchronously and encodes every retained image surface.</summary>
    public static async Task<IReadOnlyList<OfficeImageExportResult>> RenderImagesAsync(
        this HtmlConversionDocument document,
        HtmlRenderRequest request,
        CancellationToken cancellationToken = default) {
        HtmlRenderOptions resolved = HtmlRenderEngine.PrepareOptions(document, request);
        return await HtmlRenderEngine.ExecuteWithDeadlineAsync(resolved, cancellationToken, async operationCancellationToken => {
            HtmlRenderResult rendered = await HtmlRenderEngine.ExecuteCoreAsync(
                document, request, resolved, operationCancellationToken).ConfigureAwait(false);
            var images = new List<OfficeImageExportResult>(rendered.Document.Pages.Count);
            HtmlRenderResultImageExtensions.ExportImagesCore(
                rendered, MapEncoderFormat(request.Encoder), resolved, images.Add, operationCancellationToken);
            return (IReadOnlyList<OfficeImageExportResult>)images.AsReadOnly();
        }).ConfigureAwait(false);
    }

    private static HtmlRenderEncoder MapEncoder(OfficeImageExportFormat format) => format switch {
        OfficeImageExportFormat.Png => HtmlRenderEncoder.Png,
        OfficeImageExportFormat.Jpeg => HtmlRenderEncoder.Jpeg,
        OfficeImageExportFormat.Tiff => HtmlRenderEncoder.Tiff,
        OfficeImageExportFormat.Webp => HtmlRenderEncoder.Webp,
        OfficeImageExportFormat.Svg => HtmlRenderEncoder.Svg,
        _ => throw new ArgumentOutOfRangeException(nameof(format), format, "Unsupported HTML image export format.")
    };

    private static OfficeImageExportFormat MapEncoderFormat(HtmlRenderEncoder encoder) =>
        HtmlRenderResultImageExtensions.ResolveImageFormat(encoder);

}
