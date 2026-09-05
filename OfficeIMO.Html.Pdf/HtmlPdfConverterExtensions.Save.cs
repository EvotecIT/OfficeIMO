using System.Threading;
using System.Threading.Tasks;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

public static partial class HtmlPdfConverterExtensions {
    /// <summary>Converts a parsed HTML document and saves it as a PDF file.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdf(this HtmlConversionDocument document, string path, HtmlToPdfOptions? options = null, System.Threading.CancellationToken cancellationToken = default) =>
        document.ToPdfDocumentResult(options, cancellationToken).Save(path, cancellationToken);

    /// <summary>Converts a parsed HTML document and writes it as PDF to a caller-owned stream.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdf(this HtmlConversionDocument document, Stream pdfStream, HtmlToPdfOptions? options = null, System.Threading.CancellationToken cancellationToken = default) =>
        document.ToPdfDocumentResult(options, cancellationToken).Save(pdfStream);

    /// <summary>Asynchronously converts a parsed HTML document and saves it as a PDF file.</summary>
    public static async Task<PdfCore.PdfSaveResult> SaveAsPdfAsync(
        this HtmlConversionDocument document,
        string path,
        HtmlToPdfOptions? options = null,
        CancellationToken cancellationToken = default) =>
        await (await document.ToPdfDocumentResultAsync(options, cancellationToken).ConfigureAwait(false))
            .SaveAsync(path, cancellationToken).ConfigureAwait(false);

    /// <summary>Asynchronously converts a parsed HTML document and writes it as PDF to a caller-owned stream.</summary>
    public static async Task<PdfCore.PdfSaveResult> SaveAsPdfAsync(
        this HtmlConversionDocument document,
        Stream pdfStream,
        HtmlToPdfOptions? options = null,
        CancellationToken cancellationToken = default) =>
        await (await document.ToPdfDocumentResultAsync(options, cancellationToken).ConfigureAwait(false))
            .SaveAsync(pdfStream, cancellationToken).ConfigureAwait(false);

    /// <summary>Attempts to convert a parsed HTML document and save it as a PDF file.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdfResult(this HtmlConversionDocument document, string path, HtmlToPdfOptions? options = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try {
            return document.ToPdfDocumentResult(options, cancellationToken).SaveResult(path, cancellationToken);
        } catch (OperationCanceledException) { throw; }
        catch (Exception exception) {
            return PdfCore.PdfSaveResult.FromFailure(path, exception);
        }
    }

    /// <summary>Attempts to convert a parsed HTML document and write it as PDF to a caller-owned stream.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdfResult(this HtmlConversionDocument document, Stream pdfStream, HtmlToPdfOptions? options = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try {
            return document.ToPdfDocumentResult(options, cancellationToken).SaveResult(pdfStream);
        } catch (OperationCanceledException) { throw; }
        catch (Exception exception) {
            return PdfCore.PdfSaveResult.FromFailure(null, exception);
        }
    }

    /// <summary>Asynchronously attempts to convert a parsed HTML document and save it as a PDF file.</summary>
    public static async Task<PdfCore.PdfSaveResult> SaveAsPdfResultAsync(
        this HtmlConversionDocument document,
        string path,
        HtmlToPdfOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try {
            PdfCore.PdfDocumentConversionResult result = await document.ToPdfDocumentResultAsync(options, cancellationToken).ConfigureAwait(false);
            return await result.SaveResultAsync(path, cancellationToken).ConfigureAwait(false);
        } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
            throw;
        } catch (Exception exception) {
            return PdfCore.PdfSaveResult.FromFailure(path, exception);
        }
    }

    /// <summary>Asynchronously attempts to convert a parsed HTML document and write it as PDF to a caller-owned stream.</summary>
    public static async Task<PdfCore.PdfSaveResult> SaveAsPdfResultAsync(
        this HtmlConversionDocument document,
        Stream pdfStream,
        HtmlToPdfOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try {
            PdfCore.PdfDocumentConversionResult result = await document.ToPdfDocumentResultAsync(options, cancellationToken).ConfigureAwait(false);
            return await result.SaveResultAsync(pdfStream, cancellationToken).ConfigureAwait(false);
        } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
            throw;
        } catch (Exception exception) {
            return PdfCore.PdfSaveResult.FromFailure(null, exception);
        }
    }
}
