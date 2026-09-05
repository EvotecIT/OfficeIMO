using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Rtf.Pdf;

/// <summary>Converts parsed <see cref="RtfDocument"/> models to PDF.</summary>
public static partial class RtfPdfConverterExtensions {
    /// <summary>Converts an RTF document to a first-party PDF document model.</summary>
    public static PdfCore.PdfDocument ToPdfDocument(
        this RtfDocument document,
        RtfToPdfOptions? options = null) => document.ToPdfDocumentResult(options).Value;

    /// <summary>Converts an RTF document to PDF with operation-scoped diagnostics.</summary>
    public static PdfCore.PdfDocumentConversionResult ToPdfDocumentResult(
        this RtfDocument document,
        RtfToPdfOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        RtfToPdfOptions operation = (options ?? new RtfToPdfOptions()).CloneForConversion();
        PdfCore.PdfDocument pdf = RtfPdfConverter.Convert(document, operation);
        return new PdfCore.PdfDocumentConversionResult(pdf, operation.Report);
    }

    /// <summary>Converts an RTF document to PDF bytes.</summary>
    public static byte[] ToPdfBytes(this RtfDocument document, RtfToPdfOptions? options = null) =>
        document.ToPdfDocumentResult(options).ToBytes();

    /// <summary>Saves an RTF document as PDF at the specified path.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdf(this RtfDocument document, string path, RtfToPdfOptions? options = null) =>
        document.ToPdfDocumentResult(options).Save(path);

    /// <summary>Saves an RTF document as PDF to a caller-owned stream.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdf(this RtfDocument document, Stream stream, RtfToPdfOptions? options = null) =>
        document.ToPdfDocumentResult(options).Save(stream);

    /// <summary>Attempts to save an RTF document as PDF at the specified path.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdfResult(
        this RtfDocument document,
        string path,
        RtfToPdfOptions? options = null) {
        try {
            return document.ToPdfDocumentResult(options).SaveResult(path);
        } catch (Exception ex) {
            return PdfCore.PdfSaveResult.FromFailure(path, ex);
        }
    }

    /// <summary>Attempts to save an RTF document as PDF to a caller-owned stream.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdfResult(
        this RtfDocument document,
        Stream stream,
        RtfToPdfOptions? options = null) {
        try {
            return document.ToPdfDocumentResult(options).SaveResult(stream);
        } catch (Exception ex) {
            return PdfCore.PdfSaveResult.FromFailure(outputPath: null, ex);
        }
    }

    /// <summary>Converts synchronously, then asynchronously saves an RTF PDF at the specified path.</summary>
    public static Task<PdfCore.PdfSaveResult> SaveAsPdfAsync(
        this RtfDocument document,
        string path,
        RtfToPdfOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        return document.ToPdfDocumentResult(options).SaveAsync(path, cancellationToken);
    }

    /// <summary>Converts synchronously, then asynchronously saves an RTF PDF to a caller-owned stream.</summary>
    public static Task<PdfCore.PdfSaveResult> SaveAsPdfAsync(
        this RtfDocument document,
        Stream stream,
        RtfToPdfOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        return document.ToPdfDocumentResult(options).SaveAsync(stream, cancellationToken);
    }

    /// <summary>Attempts to save an RTF document as PDF at the specified path asynchronously.</summary>
    public static async Task<PdfCore.PdfSaveResult> SaveAsPdfResultAsync(
        this RtfDocument document,
        string path,
        RtfToPdfOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try {
            return await document.ToPdfDocumentResult(options)
                .SaveResultAsync(path, cancellationToken)
                .ConfigureAwait(false);
        } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
            throw;
        } catch (Exception ex) {
            return PdfCore.PdfSaveResult.FromFailure(path, ex);
        }
    }

    /// <summary>Attempts to save an RTF document as PDF to a caller-owned stream asynchronously.</summary>
    public static async Task<PdfCore.PdfSaveResult> SaveAsPdfResultAsync(
        this RtfDocument document,
        Stream stream,
        RtfToPdfOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try {
            return await document.ToPdfDocumentResult(options)
                .SaveResultAsync(stream, cancellationToken)
                .ConfigureAwait(false);
        } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
            throw;
        } catch (Exception ex) {
            return PdfCore.PdfSaveResult.FromFailure(outputPath: null, ex);
        }
    }
}
