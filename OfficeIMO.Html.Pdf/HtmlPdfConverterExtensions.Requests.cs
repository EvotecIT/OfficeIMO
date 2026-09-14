using PdfCore = OfficeIMO.Pdf;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Html.Pdf;

/// <summary>Explicit render-request HTML-to-PDF entry points.</summary>
public static partial class HtmlPdfConverterExtensions {
    /// <summary>Executes an explicit PDF render request and serializes the resulting document.</summary>
    public static byte[] RenderToPdfBytes(
        this HtmlConversionDocument document,
        HtmlRenderRequest request,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (request == null) throw new ArgumentNullException(nameof(request));
        return HtmlPdfRenderedConverter.ConvertToBytes(
            document, request, ResolvePdfOptions(request), cancellationToken);
    }

    /// <summary>Executes an explicit PDF render request and asynchronously serializes the result.</summary>
    public static async Task<byte[]> RenderToPdfBytesAsync(
        this HtmlConversionDocument document,
        HtmlRenderRequest request,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (request == null) throw new ArgumentNullException(nameof(request));
        return await HtmlPdfRenderedConverter.ConvertToBytesAsync(
            document, request, ResolvePdfOptions(request), cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Executes an explicit request and returns its retained surfaces together with the PDF output.</summary>
    public static HtmlPdfRenderRequestResult RenderToPdfResult(
        this HtmlConversionDocument document,
        HtmlRenderRequest request,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (request == null) throw new ArgumentNullException(nameof(request));
        HtmlPdfRenderResult rendered = HtmlPdfRenderedConverter.Convert(
            document, request, ResolvePdfOptions(request), cancellationToken);
        return CreateRequestResult(rendered);
    }

    /// <summary>Asynchronously executes an explicit request and returns its retained surfaces with the PDF output.</summary>
    public static async Task<HtmlPdfRenderRequestResult> RenderToPdfResultAsync(
        this HtmlConversionDocument document,
        HtmlRenderRequest request,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (request == null) throw new ArgumentNullException(nameof(request));
        HtmlPdfRenderResult rendered = await HtmlPdfRenderedConverter.ConvertAsync(
            document, request, ResolvePdfOptions(request), cancellationToken).ConfigureAwait(false);
        return CreateRequestResult(rendered);
    }

    /// <summary>Executes an explicit PDF render request and returns the PDF plus its conversion report.</summary>
    public static PdfCore.PdfDocumentConversionResult RenderToPdfDocumentResult(
        this HtmlConversionDocument document,
        HtmlRenderRequest request,
        CancellationToken cancellationToken = default) {
        return RenderToPdfResult(document, request, cancellationToken).Output;
    }

    /// <summary>Asynchronously executes an explicit PDF render request and returns the PDF plus its conversion report.</summary>
    public static async Task<PdfCore.PdfDocumentConversionResult> RenderToPdfDocumentResultAsync(
        this HtmlConversionDocument document,
        HtmlRenderRequest request,
        CancellationToken cancellationToken = default) {
        return (await RenderToPdfResultAsync(document, request, cancellationToken).ConfigureAwait(false)).Output;
    }

    private static HtmlToPdfOptions ResolvePdfOptions(HtmlRenderRequest request) {
        HtmlRenderOptions settings = request.Options;
        return settings is HtmlToPdfOptions pdf ? pdf.ClonePdf() : new HtmlToPdfOptions(settings);
    }

    private static HtmlPdfRenderRequestResult CreateRequestResult(HtmlPdfRenderResult rendered) {
        if (rendered.RenderResult == null) {
            throw new InvalidOperationException("The explicit PDF render did not retain its resolved request result.");
        }
        return new HtmlPdfRenderRequestResult(rendered.RenderResult, CreateResult(rendered));
    }
}
