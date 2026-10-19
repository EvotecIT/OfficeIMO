using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

/// <summary>Encodes retained HTML surfaces through their canonical PDF page scenes.</summary>
public static class HtmlRenderResultImageExtensions {
    /// <summary>Encodes every retained surface using the request's explicit image encoder.</summary>
    public static IReadOnlyList<OfficeImageExportResult> ExportImages(
        this HtmlRenderResult result,
        CancellationToken cancellationToken = default) {
        if (result == null) throw new ArgumentNullException(nameof(result));
        cancellationToken.ThrowIfCancellationRequested();
        OfficeImageExportFormat format = ResolveFormat(result.Request.Encoder);
        HtmlRenderOptions options = result.ResolvedOptions.Clone();
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
        var conversion = CreatePdf(result, cancellationToken);
        var pages = global::OfficeIMO.Pdf.PdfReadDocument.Open(conversion.ToBytes(cancellationToken),
            options: null, cancellationToken: cancellationToken);
        int index = 0;
        global::OfficeIMO.Pdf.PdfImageExportEngine.ExportEach(pages, format, options.ToPdfImageOptions(),
            selection: null, image => consumer(WithHtmlSource(image, result.Document.Pages[index++].PageNumber)),
            global::OfficeIMO.Pdf.PdfImageExportEngine.MapConversionDiagnostics(conversion), cancellationToken);
    }

    internal static async Task ExportImagesCoreAsync(HtmlRenderResult result, OfficeImageExportFormat format,
        HtmlRenderOptions options, OfficeImageExportAsyncConsumer consumer, CancellationToken cancellationToken) {
        var conversion = CreatePdf(result, cancellationToken);
        var pages = global::OfficeIMO.Pdf.PdfReadDocument.Open(conversion.ToBytes(cancellationToken),
            options: null, cancellationToken: cancellationToken);
        int index = 0;
        await global::OfficeIMO.Pdf.PdfImageExportEngine.ExportEachAsync(pages, format, options.ToPdfImageOptions(),
            selection: null, (image, token) => consumer(WithHtmlSource(image, result.Document.Pages[index++].PageNumber), token),
            global::OfficeIMO.Pdf.PdfImageExportEngine.MapConversionDiagnostics(conversion), cancellationToken).ConfigureAwait(false);
    }

    private static global::OfficeIMO.Pdf.PdfDocumentConversionResult CreatePdf(HtmlRenderResult result,
        CancellationToken cancellationToken) => Pdf.HtmlPdfConverterExtensions.CreateResult(
            Pdf.HtmlPdfRenderedConverter.CreatePdf(result.Document, new Pdf.HtmlToPdfOptions(result.ResolvedOptions), cancellationToken));

    private static OfficeImageExportResult WithHtmlSource(OfficeImageExportResult image, int pageNumber) =>
        image.WithIdentity("Page " + pageNumber, "HTML render page " + pageNumber);

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
