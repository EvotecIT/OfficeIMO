using System.Threading;
using System.Threading.Tasks;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

public static partial class HtmlPdfConverterExtensions {
    /// <summary>Converts a parsed HTML document and saves it as a PDF file.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdf(this HtmlConversionDocument document, string path, HtmlPdfSaveOptions? options = null) =>
        document.ToPdfDocumentResult(options).Save(path);

    /// <summary>Converts a parsed HTML document and writes it as PDF to a caller-owned stream.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdf(this HtmlConversionDocument document, Stream pdfStream, HtmlPdfSaveOptions? options = null) =>
        document.ToPdfDocumentResult(options).Save(pdfStream);

    /// <summary>Asynchronously converts a parsed HTML document and saves it as a PDF file.</summary>
    public static async Task<PdfCore.PdfSaveResult> SaveAsPdfAsync(
        this HtmlConversionDocument document,
        string path,
        HtmlPdfSaveOptions? options = null,
        CancellationToken cancellationToken = default) =>
        await (await document.ToPdfDocumentResultAsync(options, cancellationToken).ConfigureAwait(false))
            .SaveAsync(path, cancellationToken).ConfigureAwait(false);

    /// <summary>Asynchronously converts a parsed HTML document and writes it as PDF to a caller-owned stream.</summary>
    public static async Task<PdfCore.PdfSaveResult> SaveAsPdfAsync(
        this HtmlConversionDocument document,
        Stream pdfStream,
        HtmlPdfSaveOptions? options = null,
        CancellationToken cancellationToken = default) =>
        await (await document.ToPdfDocumentResultAsync(options, cancellationToken).ConfigureAwait(false))
            .SaveAsync(pdfStream, cancellationToken).ConfigureAwait(false);

    /// <summary>Attempts to convert a parsed HTML document and save it as a PDF file.</summary>
    public static PdfCore.PdfSaveResult TrySaveAsPdf(this HtmlConversionDocument document, string path, HtmlPdfSaveOptions? options = null) {
        try {
            return document.ToPdfDocumentResult(options).TrySave(path);
        } catch (Exception exception) {
            return PdfCore.PdfSaveResult.FromFailure(path, exception);
        }
    }

    /// <summary>Attempts to convert a parsed HTML document and write it as PDF to a caller-owned stream.</summary>
    public static PdfCore.PdfSaveResult TrySaveAsPdf(this HtmlConversionDocument document, Stream pdfStream, HtmlPdfSaveOptions? options = null) {
        try {
            return document.ToPdfDocumentResult(options).TrySave(pdfStream);
        } catch (Exception exception) {
            return PdfCore.PdfSaveResult.FromFailure(null, exception);
        }
    }

    /// <summary>Asynchronously attempts to convert a parsed HTML document and save it as a PDF file.</summary>
    public static async Task<PdfCore.PdfSaveResult> TrySaveAsPdfAsync(
        this HtmlConversionDocument document,
        string path,
        HtmlPdfSaveOptions? options = null,
        CancellationToken cancellationToken = default) {
        try {
            PdfCore.PdfDocumentConversionResult result = await document.ToPdfDocumentResultAsync(options, cancellationToken).ConfigureAwait(false);
            return await result.TrySaveAsync(path, cancellationToken).ConfigureAwait(false);
        } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
            throw;
        } catch (Exception exception) {
            return PdfCore.PdfSaveResult.FromFailure(path, exception);
        }
    }

    /// <summary>Asynchronously attempts to convert a parsed HTML document and write it as PDF to a caller-owned stream.</summary>
    public static async Task<PdfCore.PdfSaveResult> TrySaveAsPdfAsync(
        this HtmlConversionDocument document,
        Stream pdfStream,
        HtmlPdfSaveOptions? options = null,
        CancellationToken cancellationToken = default) {
        try {
            PdfCore.PdfDocumentConversionResult result = await document.ToPdfDocumentResultAsync(options, cancellationToken).ConfigureAwait(false);
            return await result.TrySaveAsync(pdfStream, cancellationToken).ConfigureAwait(false);
        } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
            throw;
        } catch (Exception exception) {
            return PdfCore.PdfSaveResult.FromFailure(null, exception);
        }
    }
}
