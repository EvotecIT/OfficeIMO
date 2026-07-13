using System.Threading;
using System.Threading.Tasks;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

public static partial class HtmlPdfConverterExtensions {
    /// <summary>Saves HTML as a PDF file.</summary>
    /// <example><code>html.SaveAsPdf("report.pdf");</code></example>
    public static void SaveAsPdf(this string html, string path, HtmlPdfSaveOptions? options = null) => html.ToPdfDocumentResult(options).Save(path);

    /// <summary>Saves a shared HTML conversion document as a PDF file.</summary>
    public static void SaveAsPdf(this HtmlConversionDocument document, string path, HtmlPdfSaveOptions? options = null) => document.ToPdfDocumentResult(options).Save(path);

    /// <summary>Reads UTF-8 HTML from a stream and saves it as a PDF file.</summary>
    public static void SaveAsPdf(this Stream htmlStream, string path, HtmlPdfSaveOptions? options = null) => htmlStream.ToPdfDocumentResult(options).Save(path);

    /// <summary>Writes HTML as PDF to a stream.</summary>
    public static void SaveAsPdf(this string html, Stream pdfStream, HtmlPdfSaveOptions? options = null) => html.ToPdfDocumentResult(options).Save(pdfStream);

    /// <summary>Writes a shared HTML conversion document as PDF to a stream.</summary>
    public static void SaveAsPdf(this HtmlConversionDocument document, Stream pdfStream, HtmlPdfSaveOptions? options = null) => document.ToPdfDocumentResult(options).Save(pdfStream);

    /// <summary>Reads UTF-8 HTML and writes PDF to another stream.</summary>
    public static void SaveAsPdf(this Stream htmlStream, Stream pdfStream, HtmlPdfSaveOptions? options = null) => htmlStream.ToPdfDocumentResult(options).Save(pdfStream);

    /// <summary>Asynchronously saves HTML as a PDF file.</summary>
    public static async Task SaveAsPdfAsync(this string html, string path, HtmlPdfSaveOptions? options = null, CancellationToken cancellationToken = default) =>
        await (await html.ToPdfDocumentResultAsync(options, cancellationToken).ConfigureAwait(false)).SaveAsync(path, cancellationToken).ConfigureAwait(false);

    /// <summary>Asynchronously saves a shared HTML conversion document as a PDF file.</summary>
    public static async Task SaveAsPdfAsync(this HtmlConversionDocument document, string path, HtmlPdfSaveOptions? options = null, CancellationToken cancellationToken = default) =>
        await (await document.ToPdfDocumentResultAsync(options, cancellationToken).ConfigureAwait(false)).SaveAsync(path, cancellationToken).ConfigureAwait(false);

    /// <summary>Asynchronously reads UTF-8 HTML from a stream and saves it as a PDF file.</summary>
    public static async Task SaveAsPdfAsync(this Stream htmlStream, string path, HtmlPdfSaveOptions? options = null, CancellationToken cancellationToken = default) =>
        await (await htmlStream.ToPdfDocumentResultAsync(options, cancellationToken).ConfigureAwait(false)).SaveAsync(path, cancellationToken).ConfigureAwait(false);

    /// <summary>Asynchronously writes HTML as PDF to a stream.</summary>
    public static async Task SaveAsPdfAsync(this string html, Stream pdfStream, HtmlPdfSaveOptions? options = null, CancellationToken cancellationToken = default) =>
        await (await html.ToPdfDocumentResultAsync(options, cancellationToken).ConfigureAwait(false)).SaveAsync(pdfStream, cancellationToken).ConfigureAwait(false);

    /// <summary>Asynchronously writes a shared HTML conversion document as PDF to a stream.</summary>
    public static async Task SaveAsPdfAsync(this HtmlConversionDocument document, Stream pdfStream, HtmlPdfSaveOptions? options = null, CancellationToken cancellationToken = default) =>
        await (await document.ToPdfDocumentResultAsync(options, cancellationToken).ConfigureAwait(false)).SaveAsync(pdfStream, cancellationToken).ConfigureAwait(false);

    /// <summary>Asynchronously reads UTF-8 HTML and writes PDF to another stream.</summary>
    public static async Task SaveAsPdfAsync(this Stream htmlStream, Stream pdfStream, HtmlPdfSaveOptions? options = null, CancellationToken cancellationToken = default) =>
        await (await htmlStream.ToPdfDocumentResultAsync(options, cancellationToken).ConfigureAwait(false)).SaveAsync(pdfStream, cancellationToken).ConfigureAwait(false);

    /// <summary>Attempts to save HTML as a PDF file without throwing output failures.</summary>
    public static PdfCore.PdfSaveResult TrySaveAsPdf(this string html, string path, HtmlPdfSaveOptions? options = null) {
        try { return html.ToPdfDocumentResult(options).TrySave(path); } catch (Exception ex) { return PdfCore.PdfSaveResult.FromFailure(path, ex); }
    }

    /// <summary>Attempts to save a shared HTML conversion document as a PDF file.</summary>
    public static PdfCore.PdfSaveResult TrySaveAsPdf(this HtmlConversionDocument document, string path, HtmlPdfSaveOptions? options = null) {
        try { return document.ToPdfDocumentResult(options).TrySave(path); } catch (Exception ex) { return PdfCore.PdfSaveResult.FromFailure(path, ex); }
    }

    /// <summary>Attempts to read UTF-8 HTML and save it as a PDF file.</summary>
    public static PdfCore.PdfSaveResult TrySaveAsPdf(this Stream htmlStream, string path, HtmlPdfSaveOptions? options = null) {
        try { return htmlStream.ToPdfDocumentResult(options).TrySave(path); } catch (Exception ex) { return PdfCore.PdfSaveResult.FromFailure(path, ex); }
    }

    /// <summary>Attempts to write HTML as PDF to a stream.</summary>
    public static PdfCore.PdfSaveResult TrySaveAsPdf(this string html, Stream pdfStream, HtmlPdfSaveOptions? options = null) {
        try { return html.ToPdfDocumentResult(options).TrySave(pdfStream); } catch (Exception ex) { return PdfCore.PdfSaveResult.FromFailure(null, ex); }
    }

    /// <summary>Attempts to write a shared HTML conversion document as PDF to a stream.</summary>
    public static PdfCore.PdfSaveResult TrySaveAsPdf(this HtmlConversionDocument document, Stream pdfStream, HtmlPdfSaveOptions? options = null) {
        try { return document.ToPdfDocumentResult(options).TrySave(pdfStream); } catch (Exception ex) { return PdfCore.PdfSaveResult.FromFailure(null, ex); }
    }

    /// <summary>Attempts to read UTF-8 HTML and write PDF to another stream.</summary>
    public static PdfCore.PdfSaveResult TrySaveAsPdf(this Stream htmlStream, Stream pdfStream, HtmlPdfSaveOptions? options = null) {
        try { return htmlStream.ToPdfDocumentResult(options).TrySave(pdfStream); } catch (Exception ex) { return PdfCore.PdfSaveResult.FromFailure(null, ex); }
    }

    /// <summary>Attempts to asynchronously save HTML as a PDF file.</summary>
    public static async Task<PdfCore.PdfSaveResult> TrySaveAsPdfAsync(this string html, string path, HtmlPdfSaveOptions? options = null, CancellationToken cancellationToken = default) {
        try { return await (await html.ToPdfDocumentResultAsync(options, cancellationToken).ConfigureAwait(false)).TrySaveAsync(path, cancellationToken).ConfigureAwait(false); }
        catch (Exception ex) { return PdfCore.PdfSaveResult.FromFailure(path, ex); }
    }

    /// <summary>Attempts to asynchronously write HTML as PDF to a stream.</summary>
    public static async Task<PdfCore.PdfSaveResult> TrySaveAsPdfAsync(this string html, Stream pdfStream, HtmlPdfSaveOptions? options = null, CancellationToken cancellationToken = default) {
        try { return await (await html.ToPdfDocumentResultAsync(options, cancellationToken).ConfigureAwait(false)).TrySaveAsync(pdfStream, cancellationToken).ConfigureAwait(false); }
        catch (Exception ex) { return PdfCore.PdfSaveResult.FromFailure(null, ex); }
    }

    /// <summary>Attempts to asynchronously save a shared HTML conversion document as a PDF file.</summary>
    public static async Task<PdfCore.PdfSaveResult> TrySaveAsPdfAsync(this HtmlConversionDocument document, string path, HtmlPdfSaveOptions? options = null, CancellationToken cancellationToken = default) {
        try { return await (await document.ToPdfDocumentResultAsync(options, cancellationToken).ConfigureAwait(false)).TrySaveAsync(path, cancellationToken).ConfigureAwait(false); }
        catch (Exception ex) { return PdfCore.PdfSaveResult.FromFailure(path, ex); }
    }

    /// <summary>Attempts to asynchronously write a shared HTML conversion document as PDF to a stream.</summary>
    public static async Task<PdfCore.PdfSaveResult> TrySaveAsPdfAsync(this HtmlConversionDocument document, Stream pdfStream, HtmlPdfSaveOptions? options = null, CancellationToken cancellationToken = default) {
        try { return await (await document.ToPdfDocumentResultAsync(options, cancellationToken).ConfigureAwait(false)).TrySaveAsync(pdfStream, cancellationToken).ConfigureAwait(false); }
        catch (Exception ex) { return PdfCore.PdfSaveResult.FromFailure(null, ex); }
    }

    /// <summary>Attempts to asynchronously read UTF-8 HTML and save it as a PDF file.</summary>
    public static async Task<PdfCore.PdfSaveResult> TrySaveAsPdfAsync(this Stream htmlStream, string path, HtmlPdfSaveOptions? options = null, CancellationToken cancellationToken = default) {
        try { return await (await htmlStream.ToPdfDocumentResultAsync(options, cancellationToken).ConfigureAwait(false)).TrySaveAsync(path, cancellationToken).ConfigureAwait(false); }
        catch (Exception ex) { return PdfCore.PdfSaveResult.FromFailure(path, ex); }
    }

    /// <summary>Attempts to asynchronously read UTF-8 HTML and write PDF to another stream.</summary>
    public static async Task<PdfCore.PdfSaveResult> TrySaveAsPdfAsync(this Stream htmlStream, Stream pdfStream, HtmlPdfSaveOptions? options = null, CancellationToken cancellationToken = default) {
        try { return await (await htmlStream.ToPdfDocumentResultAsync(options, cancellationToken).ConfigureAwait(false)).TrySaveAsync(pdfStream, cancellationToken).ConfigureAwait(false); }
        catch (Exception ex) { return PdfCore.PdfSaveResult.FromFailure(null, ex); }
    }
}
