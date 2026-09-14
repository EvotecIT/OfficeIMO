using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

/// <summary>Encodes retained HTML render results through the shared Drawing image encoders.</summary>
public static class HtmlRenderResultImageExtensions {
    /// <summary>Encodes every retained surface using the request's explicit image encoder.</summary>
    public static IReadOnlyList<OfficeImageExportResult> ExportImages(
        this HtmlRenderResult result,
        CancellationToken cancellationToken = default) {
        if (result == null) throw new ArgumentNullException(nameof(result));
        cancellationToken.ThrowIfCancellationRequested();
        OfficeImageExportFormat format = ResolveFormat(result.Request.Encoder);
        HtmlRenderOptions options = result.Request.ResolveOptions();
        var images = new List<OfficeImageExportResult>(result.Document.Pages.Count);
        HtmlRenderEngine.ExecuteWithDeadline(options, cancellationToken, operationCancellationToken => {
            ExportImagesCore(result, format, options, images.Add, operationCancellationToken);
            return true;
        });
        return images.AsReadOnly();
    }

    /// <summary>Encodes the only retained surface using the request's explicit image encoder.</summary>
    public static OfficeImageExportResult ExportImage(
        this HtmlRenderResult result,
        CancellationToken cancellationToken = default) {
        if (result == null) throw new ArgumentNullException(nameof(result));
        if (result.Document.Pages.Count != 1) {
            throw new InvalidOperationException("Single-image export requires a selected, viewport, continuous, or stitched one-surface request.");
        }
        return result.ExportImages(cancellationToken)[0];
    }

    internal static void ExportImagesCore(
        HtmlRenderResult result,
        OfficeImageExportFormat format,
        HtmlRenderOptions options,
        OfficeImageExportConsumer consumer,
        CancellationToken cancellationToken) {
        OfficeImageExportBatchProcessor.ForEachOrdered(
            result.Document.Pages,
            options.MaximumDegreeOfParallelism,
            (page, _, token) => HtmlImageExportExtensions.RenderPage(
                page, format, options, result.Document.DiagnosticReport, token),
            consumer,
            cancellationToken,
            options);
    }

    internal static OfficeImageExportFormat ResolveImageFormat(HtmlRenderEncoder encoder) => ResolveFormat(encoder);

    private static OfficeImageExportFormat ResolveFormat(HtmlRenderEncoder encoder) => encoder switch {
        HtmlRenderEncoder.Png => OfficeImageExportFormat.Png,
        HtmlRenderEncoder.Jpeg => OfficeImageExportFormat.Jpeg,
        HtmlRenderEncoder.Tiff => OfficeImageExportFormat.Tiff,
        HtmlRenderEncoder.Webp => OfficeImageExportFormat.Webp,
        HtmlRenderEncoder.Svg => OfficeImageExportFormat.Svg,
        _ => throw new InvalidOperationException($"Render encoder '{encoder}' is not an image encoder.")
    };
}
