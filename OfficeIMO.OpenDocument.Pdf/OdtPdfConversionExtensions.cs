using OfficeIMO.Word.OpenDocument;
using WordPdf = OfficeIMO.Word.Pdf;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.OpenDocument.Pdf;

/// <summary>Direct, loss-aware ODT to PDF conversion through the Word semantic and PDF engines.</summary>
public static class OdtPdfConversionExtensions {
    /// <summary>Converts an ODT document to the first-party PDF document model.</summary>
    public static PdfCore.PdfDocument ToPdfDocument(
        this OdtDocument source,
        WordOpenDocumentConversionOptions? conversionOptions = null,
        WordPdf.PdfSaveOptions? pdfOptions = null) =>
        source.ToPdfDocumentResult(conversionOptions, pdfOptions).Value;

    /// <summary>Converts an ODT document to PDF and preserves diagnostics from both conversion stages.</summary>
    public static PdfCore.PdfDocumentConversionResult ToPdfDocumentResult(
        this OdtDocument source,
        WordOpenDocumentConversionOptions? conversionOptions = null,
        WordPdf.PdfSaveOptions? pdfOptions = null) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        OdfConversionResult<OfficeIMO.Word.WordDocument> conversion =
            source.ToWordDocumentResult(conversionOptions);
        using (conversion.Value) {
            PdfCore.PdfDocumentConversionResult result =
                WordPdf.WordPdfConverterExtensions.ToPdfDocumentResult(conversion.Value, pdfOptions);
            return OdfPdfConversionDiagnostics.Attach(result, conversion.Report);
        }
    }

    /// <summary>Converts an ODT document to PDF bytes.</summary>
    public static byte[] ToPdf(this OdtDocument source, WordOpenDocumentConversionOptions? conversionOptions = null, WordPdf.PdfSaveOptions? pdfOptions = null) =>
        source.ToPdfDocumentResult(conversionOptions, pdfOptions).ToBytes();

    /// <summary>Saves an ODT document as PDF.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdf(this OdtDocument source, string path, WordOpenDocumentConversionOptions? conversionOptions = null, WordPdf.PdfSaveOptions? pdfOptions = null) =>
        source.ToPdfDocumentResult(conversionOptions, pdfOptions).Save(path);

    /// <summary>Writes an ODT document as PDF to a caller-owned stream.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdf(this OdtDocument source, Stream stream, WordOpenDocumentConversionOptions? conversionOptions = null, WordPdf.PdfSaveOptions? pdfOptions = null) =>
        source.ToPdfDocumentResult(conversionOptions, pdfOptions).Save(stream);

    /// <summary>Attempts to save an ODT document as PDF and returns structured failure evidence.</summary>
    public static PdfCore.PdfSaveResult TrySaveAsPdf(this OdtDocument source, string path, WordOpenDocumentConversionOptions? conversionOptions = null, WordPdf.PdfSaveOptions? pdfOptions = null) {
        try { return source.ToPdfDocumentResult(conversionOptions, pdfOptions).TrySave(path); }
        catch (Exception ex) { return PdfCore.PdfSaveResult.FromFailure(path, ex); }
    }

    /// <summary>Attempts to write an ODT document as PDF and returns structured failure evidence.</summary>
    public static PdfCore.PdfSaveResult TrySaveAsPdf(this OdtDocument source, Stream stream, WordOpenDocumentConversionOptions? conversionOptions = null, WordPdf.PdfSaveOptions? pdfOptions = null) {
        try { return source.ToPdfDocumentResult(conversionOptions, pdfOptions).TrySave(stream); }
        catch (Exception ex) { return PdfCore.PdfSaveResult.FromFailure(outputPath: null, ex); }
    }

    /// <summary>Converts synchronously, then asynchronously saves an ODT document as PDF.</summary>
    public static Task<PdfCore.PdfSaveResult> SaveAsPdfAsync(this OdtDocument source, string path, WordOpenDocumentConversionOptions? conversionOptions = null, WordPdf.PdfSaveOptions? pdfOptions = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        return source.ToPdfDocumentResult(conversionOptions, pdfOptions).SaveAsync(path, cancellationToken);
    }

    /// <summary>Converts synchronously, then asynchronously writes an ODT document as PDF.</summary>
    public static Task<PdfCore.PdfSaveResult> SaveAsPdfAsync(this OdtDocument source, Stream stream, WordOpenDocumentConversionOptions? conversionOptions = null, WordPdf.PdfSaveOptions? pdfOptions = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        return source.ToPdfDocumentResult(conversionOptions, pdfOptions).SaveAsync(stream, cancellationToken);
    }

    /// <summary>Attempts to asynchronously save an ODT document as PDF and returns structured failure evidence.</summary>
    public static async Task<PdfCore.PdfSaveResult> TrySaveAsPdfAsync(this OdtDocument source, string path, WordOpenDocumentConversionOptions? conversionOptions = null, WordPdf.PdfSaveOptions? pdfOptions = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try { return await source.ToPdfDocumentResult(conversionOptions, pdfOptions).TrySaveAsync(path, cancellationToken).ConfigureAwait(false); }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { throw; }
        catch (Exception ex) { return PdfCore.PdfSaveResult.FromFailure(path, ex); }
    }

    /// <summary>Attempts to asynchronously write an ODT document as PDF and returns structured failure evidence.</summary>
    public static async Task<PdfCore.PdfSaveResult> TrySaveAsPdfAsync(this OdtDocument source, Stream stream, WordOpenDocumentConversionOptions? conversionOptions = null, WordPdf.PdfSaveOptions? pdfOptions = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try { return await source.ToPdfDocumentResult(conversionOptions, pdfOptions).TrySaveAsync(stream, cancellationToken).ConfigureAwait(false); }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { throw; }
        catch (Exception ex) { return PdfCore.PdfSaveResult.FromFailure(outputPath: null, ex); }
    }
}
