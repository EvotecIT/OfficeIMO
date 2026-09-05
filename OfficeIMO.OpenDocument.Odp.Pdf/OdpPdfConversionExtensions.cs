using OfficeIMO.PowerPoint.OpenDocument;
using PowerPointPdf = OfficeIMO.PowerPoint.Pdf;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.OpenDocument.Odp.Pdf;

/// <summary>Direct, loss-aware ODP to PDF conversion through the PowerPoint semantic and PDF engines.</summary>
public static class OdpPdfConversionExtensions {
    /// <summary>Converts an ODP presentation to the first-party PDF document model.</summary>
    public static PdfCore.PdfDocument ToPdfDocument(this OdpPresentation source, PowerPointOpenDocumentConversionOptions? conversionOptions = null, PowerPointPdf.PowerPointToPdfOptions? pdfOptions = null) =>
        source.ToPdfDocumentResult(conversionOptions, pdfOptions).Value;

    /// <summary>Converts an ODP presentation to PDF and preserves diagnostics from both conversion stages.</summary>
    public static PdfCore.PdfDocumentConversionResult ToPdfDocumentResult(this OdpPresentation source, PowerPointOpenDocumentConversionOptions? conversionOptions = null, PowerPointPdf.PowerPointToPdfOptions? pdfOptions = null) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        OdfConversionResult<OfficeIMO.PowerPoint.PowerPointPresentation> conversion = source.ToPowerPointPresentationResult(conversionOptions);
        using (conversion.Value) {
            PdfCore.PdfDocumentConversionResult result = PowerPointPdf.PowerPointPdfConverterExtensions.ToPdfDocumentResult(conversion.Value, pdfOptions);
            return result.WithSourceConversionReport(conversion.Report);
        }
    }

    /// <summary>Converts an ODP presentation to PDF bytes.</summary>
    public static byte[] ToPdfBytes(this OdpPresentation source, PowerPointOpenDocumentConversionOptions? conversionOptions = null, PowerPointPdf.PowerPointToPdfOptions? pdfOptions = null) =>
        source.ToPdfDocumentResult(conversionOptions, pdfOptions).ToBytes();

    /// <summary>Saves an ODP presentation as PDF.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdf(this OdpPresentation source, string path, PowerPointOpenDocumentConversionOptions? conversionOptions = null, PowerPointPdf.PowerPointToPdfOptions? pdfOptions = null) =>
        source.ToPdfDocumentResult(conversionOptions, pdfOptions).Save(path);

    /// <summary>Writes an ODP presentation as PDF to a caller-owned stream.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdf(this OdpPresentation source, Stream stream, PowerPointOpenDocumentConversionOptions? conversionOptions = null, PowerPointPdf.PowerPointToPdfOptions? pdfOptions = null) =>
        source.ToPdfDocumentResult(conversionOptions, pdfOptions).Save(stream);

    /// <summary>Attempts to save an ODP presentation as PDF and returns structured failure evidence.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdfResult(this OdpPresentation source, string path, PowerPointOpenDocumentConversionOptions? conversionOptions = null, PowerPointPdf.PowerPointToPdfOptions? pdfOptions = null) {
        try { return source.ToPdfDocumentResult(conversionOptions, pdfOptions).SaveResult(path); }
        catch (Exception ex) { return PdfCore.PdfSaveResult.FromFailure(path, ex); }
    }

    /// <summary>Attempts to write an ODP presentation as PDF and returns structured failure evidence.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdfResult(this OdpPresentation source, Stream stream, PowerPointOpenDocumentConversionOptions? conversionOptions = null, PowerPointPdf.PowerPointToPdfOptions? pdfOptions = null) {
        try { return source.ToPdfDocumentResult(conversionOptions, pdfOptions).SaveResult(stream); }
        catch (Exception ex) { return PdfCore.PdfSaveResult.FromFailure(outputPath: null, ex); }
    }

    /// <summary>Converts synchronously, then asynchronously saves an ODP presentation as PDF.</summary>
    public static Task<PdfCore.PdfSaveResult> SaveAsPdfAsync(this OdpPresentation source, string path, PowerPointOpenDocumentConversionOptions? conversionOptions = null, PowerPointPdf.PowerPointToPdfOptions? pdfOptions = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        return source.ToPdfDocumentResult(conversionOptions, pdfOptions).SaveAsync(path, cancellationToken);
    }

    /// <summary>Converts synchronously, then asynchronously writes an ODP presentation as PDF.</summary>
    public static Task<PdfCore.PdfSaveResult> SaveAsPdfAsync(this OdpPresentation source, Stream stream, PowerPointOpenDocumentConversionOptions? conversionOptions = null, PowerPointPdf.PowerPointToPdfOptions? pdfOptions = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        return source.ToPdfDocumentResult(conversionOptions, pdfOptions).SaveAsync(stream, cancellationToken);
    }

    /// <summary>Attempts to asynchronously save an ODP presentation as PDF and returns structured failure evidence.</summary>
    public static async Task<PdfCore.PdfSaveResult> SaveAsPdfResultAsync(this OdpPresentation source, string path, PowerPointOpenDocumentConversionOptions? conversionOptions = null, PowerPointPdf.PowerPointToPdfOptions? pdfOptions = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try { return await source.ToPdfDocumentResult(conversionOptions, pdfOptions).SaveResultAsync(path, cancellationToken).ConfigureAwait(false); }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { throw; }
        catch (Exception ex) { return PdfCore.PdfSaveResult.FromFailure(path, ex); }
    }

    /// <summary>Attempts to asynchronously write an ODP presentation as PDF and returns structured failure evidence.</summary>
    public static async Task<PdfCore.PdfSaveResult> SaveAsPdfResultAsync(this OdpPresentation source, Stream stream, PowerPointOpenDocumentConversionOptions? conversionOptions = null, PowerPointPdf.PowerPointToPdfOptions? pdfOptions = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try { return await source.ToPdfDocumentResult(conversionOptions, pdfOptions).SaveResultAsync(stream, cancellationToken).ConfigureAwait(false); }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { throw; }
        catch (Exception ex) { return PdfCore.PdfSaveResult.FromFailure(outputPath: null, ex); }
    }

    /// <summary>Reconstructs an ODP presentation from an opened PDF.</summary>
    public static OdpPresentation ToOdpPresentation(this PdfCore.PdfDocument source, PowerPointPdf.PdfPowerPointImportOptions? pdfOptions = null, PowerPointOpenDocumentConversionOptions? openDocumentOptions = null) =>
        source.ToOdpPresentationResult(pdfOptions, openDocumentOptions).Value;

    /// <summary>Reconstructs an ODP presentation and preserves diagnostics from both conversion stages.</summary>
    public static PdfOdpConversionResult ToOdpPresentationResult(this PdfCore.PdfDocument source, PowerPointPdf.PdfPowerPointImportOptions? pdfOptions = null, PowerPointOpenDocumentConversionOptions? openDocumentOptions = null) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        PowerPointPdf.PdfPowerPointConversionResult pdf =
            PowerPointPdf.PowerPointPdfConverterExtensions.ToPowerPointPresentationResult(source, pdfOptions);
        using (pdf.Value) {
            OdfConversionResult<OdpPresentation> odp = pdf.Value.ToOpenDocumentResult(openDocumentOptions);
            return new PdfOdpConversionResult(
                odp.Value,
                new PdfOdpConversionReport(pdf.Report, odp.Report));
        }
    }

    /// <summary>Reconstructs an ODP presentation from an already loaded logical PDF model.</summary>
    /// <remarks>A logical PDF model resolves Auto to editable tables and also supports explicit editable-content reconstruction. Visual and hybrid projections require an opened PDF.</remarks>
    public static OdpPresentation ToOdpPresentation(this PdfCore.PdfDocumentReadResult source, PowerPointPdf.PdfPowerPointImportOptions? pdfOptions = null, PowerPointOpenDocumentConversionOptions? openDocumentOptions = null) =>
        source.ToOdpPresentationResult(pdfOptions, openDocumentOptions).Value;

    /// <summary>Reconstructs an ODP presentation from a logical PDF model and preserves both stage reports.</summary>
    /// <remarks>A logical PDF model resolves Auto to editable tables and also supports explicit editable-content reconstruction. Visual and hybrid projections require an opened PDF.</remarks>
    public static PdfOdpConversionResult ToOdpPresentationResult(this PdfCore.PdfDocumentReadResult source, PowerPointPdf.PdfPowerPointImportOptions? pdfOptions = null, PowerPointOpenDocumentConversionOptions? openDocumentOptions = null) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        PowerPointPdf.PdfPowerPointConversionResult pdf =
            PowerPointPdf.PowerPointPdfConverterExtensions.ToPowerPointPresentationResult(source, pdfOptions);
        using (pdf.Value) {
            OdfConversionResult<OdpPresentation> odp = pdf.Value.ToOpenDocumentResult(openDocumentOptions);
            return new PdfOdpConversionResult(
                odp.Value,
                new PdfOdpConversionReport(pdf.Report, odp.Report));
        }
    }

    /// <summary>Reconstructs and saves an ODP presentation from an opened PDF.</summary>
    public static PdfOdpConversionReport SaveAsOdp(this PdfCore.PdfDocument source, string path, PowerPointPdf.PdfPowerPointImportOptions? pdfOptions = null, PowerPointOpenDocumentConversionOptions? openDocumentOptions = null) {
        PdfOdpConversionResult result = source.ToOdpPresentationResult(pdfOptions, openDocumentOptions);
        result.Value.Save(path);
        return result.Report;
    }

    /// <summary>Reconstructs and writes an ODP presentation from an opened PDF.</summary>
    public static PdfOdpConversionReport SaveAsOdp(this PdfCore.PdfDocument source, Stream stream, PowerPointPdf.PdfPowerPointImportOptions? pdfOptions = null, PowerPointOpenDocumentConversionOptions? openDocumentOptions = null) {
        PdfOdpConversionResult result = source.ToOdpPresentationResult(pdfOptions, openDocumentOptions);
        result.Value.Save(stream);
        return result.Report;
    }

    /// <summary>Reconstructs and saves an ODP presentation from a logical PDF model.</summary>
    public static PdfOdpConversionReport SaveAsOdp(this PdfCore.PdfDocumentReadResult source, string path, PowerPointPdf.PdfPowerPointImportOptions? pdfOptions = null, PowerPointOpenDocumentConversionOptions? openDocumentOptions = null) {
        PdfOdpConversionResult result = source.ToOdpPresentationResult(pdfOptions, openDocumentOptions);
        result.Value.Save(path);
        return result.Report;
    }

    /// <summary>Reconstructs and writes an ODP presentation from a logical PDF model.</summary>
    public static PdfOdpConversionReport SaveAsOdp(this PdfCore.PdfDocumentReadResult source, Stream stream, PowerPointPdf.PdfPowerPointImportOptions? pdfOptions = null, PowerPointOpenDocumentConversionOptions? openDocumentOptions = null) {
        PdfOdpConversionResult result = source.ToOdpPresentationResult(pdfOptions, openDocumentOptions);
        result.Value.Save(stream);
        return result.Report;
    }

    /// <summary>Reconstructs synchronously, then asynchronously saves an ODP presentation from an opened PDF.</summary>
    public static async Task<PdfOdpConversionReport> SaveAsOdpAsync(this PdfCore.PdfDocument source, string path, PowerPointPdf.PdfPowerPointImportOptions? pdfOptions = null, PowerPointOpenDocumentConversionOptions? openDocumentOptions = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfOdpConversionResult result = source.ToOdpPresentationResult(pdfOptions, openDocumentOptions);
        await result.Value.SaveAsync(path, cancellationToken: cancellationToken).ConfigureAwait(false);
        return result.Report;
    }

    /// <summary>Reconstructs synchronously, then asynchronously writes an ODP presentation from an opened PDF.</summary>
    public static async Task<PdfOdpConversionReport> SaveAsOdpAsync(this PdfCore.PdfDocument source, Stream stream, PowerPointPdf.PdfPowerPointImportOptions? pdfOptions = null, PowerPointOpenDocumentConversionOptions? openDocumentOptions = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfOdpConversionResult result = source.ToOdpPresentationResult(pdfOptions, openDocumentOptions);
        await result.Value.SaveAsync(stream, cancellationToken: cancellationToken).ConfigureAwait(false);
        return result.Report;
    }

    /// <summary>Reconstructs synchronously, then asynchronously saves an ODP presentation from a logical PDF model.</summary>
    public static async Task<PdfOdpConversionReport> SaveAsOdpAsync(this PdfCore.PdfDocumentReadResult source, string path, PowerPointPdf.PdfPowerPointImportOptions? pdfOptions = null, PowerPointOpenDocumentConversionOptions? openDocumentOptions = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfOdpConversionResult result = source.ToOdpPresentationResult(pdfOptions, openDocumentOptions);
        await result.Value.SaveAsync(path, cancellationToken: cancellationToken).ConfigureAwait(false);
        return result.Report;
    }

    /// <summary>Reconstructs synchronously, then asynchronously writes an ODP presentation from a logical PDF model.</summary>
    public static async Task<PdfOdpConversionReport> SaveAsOdpAsync(this PdfCore.PdfDocumentReadResult source, Stream stream, PowerPointPdf.PdfPowerPointImportOptions? pdfOptions = null, PowerPointOpenDocumentConversionOptions? openDocumentOptions = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfOdpConversionResult result = source.ToOdpPresentationResult(pdfOptions, openDocumentOptions);
        await result.Value.SaveAsync(stream, cancellationToken: cancellationToken).ConfigureAwait(false);
        return result.Report;
    }
}
