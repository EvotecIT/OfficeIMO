using OfficeIMO.PowerPoint.OpenDocument;
using PowerPointPdf = OfficeIMO.PowerPoint.Pdf;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.OpenDocument.Pdf;

/// <summary>Direct, loss-aware ODP to PDF conversion through the PowerPoint semantic and PDF engines.</summary>
public static class OdpPdfConversionExtensions {
    /// <summary>Converts an ODP presentation to the first-party PDF document model.</summary>
    public static PdfCore.PdfDocument ToPdfDocument(this OdpPresentation source, PowerPointOpenDocumentConversionOptions? conversionOptions = null, PowerPointPdf.PowerPointPdfSaveOptions? pdfOptions = null) =>
        source.ToPdfDocumentResult(conversionOptions, pdfOptions).Value;

    /// <summary>Converts an ODP presentation to PDF and preserves diagnostics from both conversion stages.</summary>
    public static PdfCore.PdfDocumentConversionResult ToPdfDocumentResult(this OdpPresentation source, PowerPointOpenDocumentConversionOptions? conversionOptions = null, PowerPointPdf.PowerPointPdfSaveOptions? pdfOptions = null) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        OdfConversionResult<OfficeIMO.PowerPoint.PowerPointPresentation> conversion = source.ToPowerPointPresentationResult(conversionOptions);
        using (conversion.Value) {
            PdfCore.PdfDocumentConversionResult result = PowerPointPdf.PowerPointPdfConverterExtensions.ToPdfDocumentResult(conversion.Value, pdfOptions);
            return OdfPdfConversionDiagnostics.Attach(result, conversion.Report);
        }
    }

    /// <summary>Converts an ODP presentation to PDF bytes.</summary>
    public static byte[] ToPdf(this OdpPresentation source, PowerPointOpenDocumentConversionOptions? conversionOptions = null, PowerPointPdf.PowerPointPdfSaveOptions? pdfOptions = null) =>
        source.ToPdfDocumentResult(conversionOptions, pdfOptions).ToBytes();

    /// <summary>Saves an ODP presentation as PDF.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdf(this OdpPresentation source, string path, PowerPointOpenDocumentConversionOptions? conversionOptions = null, PowerPointPdf.PowerPointPdfSaveOptions? pdfOptions = null) =>
        source.ToPdfDocumentResult(conversionOptions, pdfOptions).Save(path);

    /// <summary>Writes an ODP presentation as PDF to a caller-owned stream.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdf(this OdpPresentation source, Stream stream, PowerPointOpenDocumentConversionOptions? conversionOptions = null, PowerPointPdf.PowerPointPdfSaveOptions? pdfOptions = null) =>
        source.ToPdfDocumentResult(conversionOptions, pdfOptions).Save(stream);

    /// <summary>Attempts to save an ODP presentation as PDF and returns structured failure evidence.</summary>
    public static PdfCore.PdfSaveResult TrySaveAsPdf(this OdpPresentation source, string path, PowerPointOpenDocumentConversionOptions? conversionOptions = null, PowerPointPdf.PowerPointPdfSaveOptions? pdfOptions = null) {
        try { return source.ToPdfDocumentResult(conversionOptions, pdfOptions).TrySave(path); }
        catch (Exception ex) { return PdfCore.PdfSaveResult.FromFailure(path, ex); }
    }

    /// <summary>Attempts to write an ODP presentation as PDF and returns structured failure evidence.</summary>
    public static PdfCore.PdfSaveResult TrySaveAsPdf(this OdpPresentation source, Stream stream, PowerPointOpenDocumentConversionOptions? conversionOptions = null, PowerPointPdf.PowerPointPdfSaveOptions? pdfOptions = null) {
        try { return source.ToPdfDocumentResult(conversionOptions, pdfOptions).TrySave(stream); }
        catch (Exception ex) { return PdfCore.PdfSaveResult.FromFailure(outputPath: null, ex); }
    }

    /// <summary>Converts synchronously, then asynchronously saves an ODP presentation as PDF.</summary>
    public static Task<PdfCore.PdfSaveResult> SaveAsPdfAsync(this OdpPresentation source, string path, PowerPointOpenDocumentConversionOptions? conversionOptions = null, PowerPointPdf.PowerPointPdfSaveOptions? pdfOptions = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        return source.ToPdfDocumentResult(conversionOptions, pdfOptions).SaveAsync(path, cancellationToken);
    }

    /// <summary>Converts synchronously, then asynchronously writes an ODP presentation as PDF.</summary>
    public static Task<PdfCore.PdfSaveResult> SaveAsPdfAsync(this OdpPresentation source, Stream stream, PowerPointOpenDocumentConversionOptions? conversionOptions = null, PowerPointPdf.PowerPointPdfSaveOptions? pdfOptions = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        return source.ToPdfDocumentResult(conversionOptions, pdfOptions).SaveAsync(stream, cancellationToken);
    }

    /// <summary>Attempts to asynchronously save an ODP presentation as PDF and returns structured failure evidence.</summary>
    public static async Task<PdfCore.PdfSaveResult> TrySaveAsPdfAsync(this OdpPresentation source, string path, PowerPointOpenDocumentConversionOptions? conversionOptions = null, PowerPointPdf.PowerPointPdfSaveOptions? pdfOptions = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try { return await source.ToPdfDocumentResult(conversionOptions, pdfOptions).TrySaveAsync(path, cancellationToken).ConfigureAwait(false); }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { throw; }
        catch (Exception ex) { return PdfCore.PdfSaveResult.FromFailure(path, ex); }
    }

    /// <summary>Attempts to asynchronously write an ODP presentation as PDF and returns structured failure evidence.</summary>
    public static async Task<PdfCore.PdfSaveResult> TrySaveAsPdfAsync(this OdpPresentation source, Stream stream, PowerPointOpenDocumentConversionOptions? conversionOptions = null, PowerPointPdf.PowerPointPdfSaveOptions? pdfOptions = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try { return await source.ToPdfDocumentResult(conversionOptions, pdfOptions).TrySaveAsync(stream, cancellationToken).ConfigureAwait(false); }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { throw; }
        catch (Exception ex) { return PdfCore.PdfSaveResult.FromFailure(outputPath: null, ex); }
    }
}
