using OfficeIMO.PowerPoint.OpenDocument;
using PowerPointPdf = OfficeIMO.PowerPoint.Pdf;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.OpenDocument.Odp.Pdf;

/// <summary>Direct, loss-aware ODP to PDF conversion through the PowerPoint semantic and PDF engines.</summary>
public static class OdpPdfConversionExtensions {
    /// <summary>Converts an ODP presentation to the first-party PDF document model.</summary>
    public static PdfCore.PdfDocument ToPdfDocument(this OdpPresentation source, PowerPointOpenDocumentConversionOptions? conversionOptions = null, PowerPointPdf.PowerPointToPdfOptions? pdfOptions = null, System.Threading.CancellationToken cancellationToken = default) =>
        source.ToPdfDocumentResult(conversionOptions, pdfOptions, cancellationToken).Value;

    /// <summary>Converts an ODP presentation to PDF and preserves diagnostics from both conversion stages.</summary>
    public static PdfCore.PdfDocumentConversionResult ToPdfDocumentResult(this OdpPresentation source, PowerPointOpenDocumentConversionOptions? conversionOptions = null, PowerPointPdf.PowerPointToPdfOptions? pdfOptions = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (source == null) throw new ArgumentNullException(nameof(source));
        OdfConversionResult<OfficeIMO.PowerPoint.PowerPointPresentation> conversion = source.ToPowerPointPresentationResult(conversionOptions);
        using (conversion.Value) {
            PdfCore.PdfDocumentConversionResult result = PowerPointPdf.PowerPointPdfConverterExtensions.ToPdfDocumentResult(conversion.Value, pdfOptions, cancellationToken);
            return result.WithSourceConversionReport(conversion.Report);
        }
    }

    /// <summary>Converts an ODP presentation to PDF bytes.</summary>
    public static byte[] ToPdfBytes(this OdpPresentation source, PowerPointOpenDocumentConversionOptions? conversionOptions = null, PowerPointPdf.PowerPointToPdfOptions? pdfOptions = null, System.Threading.CancellationToken cancellationToken = default) =>
        source.ToPdfDocumentResult(conversionOptions, pdfOptions, cancellationToken).ToBytes(cancellationToken);

    /// <summary>Saves an ODP presentation as PDF.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdf(this OdpPresentation source, string path, PowerPointOpenDocumentConversionOptions? conversionOptions = null, PowerPointPdf.PowerPointToPdfOptions? pdfOptions = null, System.Threading.CancellationToken cancellationToken = default) =>
        source.ToPdfDocumentResult(conversionOptions, pdfOptions, cancellationToken).Save(path, cancellationToken);

    /// <summary>Writes an ODP presentation as PDF to a caller-owned stream.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdf(this OdpPresentation source, Stream stream, PowerPointOpenDocumentConversionOptions? conversionOptions = null, PowerPointPdf.PowerPointToPdfOptions? pdfOptions = null, System.Threading.CancellationToken cancellationToken = default) =>
        source.ToPdfDocumentResult(conversionOptions, pdfOptions, cancellationToken).Save(stream, cancellationToken);

    /// <summary>Attempts to save an ODP presentation as PDF and returns structured failure evidence.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdfResult(this OdpPresentation source, string path, PowerPointOpenDocumentConversionOptions? conversionOptions = null, PowerPointPdf.PowerPointToPdfOptions? pdfOptions = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try { return source.ToPdfDocumentResult(conversionOptions, pdfOptions, cancellationToken).SaveResult(path, cancellationToken); }
        catch (OperationCanceledException) { throw; }
        catch (Exception ex) { return PdfCore.PdfSaveResult.FromFailure(path, ex); }
    }

    /// <summary>Attempts to write an ODP presentation as PDF and returns structured failure evidence.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdfResult(this OdpPresentation source, Stream stream, PowerPointOpenDocumentConversionOptions? conversionOptions = null, PowerPointPdf.PowerPointToPdfOptions? pdfOptions = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try { return source.ToPdfDocumentResult(conversionOptions, pdfOptions, cancellationToken).SaveResult(stream, cancellationToken); }
        catch (OperationCanceledException) { throw; }
        catch (Exception ex) { return PdfCore.PdfSaveResult.FromFailure(outputPath: null, ex); }
    }

    /// <summary>Converts synchronously, then asynchronously saves an ODP presentation as PDF.</summary>
    public static Task<PdfCore.PdfSaveResult> SaveAsPdfAsync(this OdpPresentation source, string path, PowerPointOpenDocumentConversionOptions? conversionOptions = null, PowerPointPdf.PowerPointToPdfOptions? pdfOptions = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        return source.ToPdfDocumentResult(conversionOptions, pdfOptions, cancellationToken).SaveAsync(path, cancellationToken);
    }

    /// <summary>Converts synchronously, then asynchronously writes an ODP presentation as PDF.</summary>
    public static Task<PdfCore.PdfSaveResult> SaveAsPdfAsync(this OdpPresentation source, Stream stream, PowerPointOpenDocumentConversionOptions? conversionOptions = null, PowerPointPdf.PowerPointToPdfOptions? pdfOptions = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        return source.ToPdfDocumentResult(conversionOptions, pdfOptions, cancellationToken).SaveAsync(stream, cancellationToken);
    }

    /// <summary>Attempts to asynchronously save an ODP presentation as PDF and returns structured failure evidence.</summary>
    public static async Task<PdfCore.PdfSaveResult> SaveAsPdfResultAsync(this OdpPresentation source, string path, PowerPointOpenDocumentConversionOptions? conversionOptions = null, PowerPointPdf.PowerPointToPdfOptions? pdfOptions = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try { return await source.ToPdfDocumentResult(conversionOptions, pdfOptions, cancellationToken).SaveResultAsync(path, cancellationToken).ConfigureAwait(false); }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { throw; }
        catch (Exception ex) { return PdfCore.PdfSaveResult.FromFailure(path, ex); }
    }

    /// <summary>Attempts to asynchronously write an ODP presentation as PDF and returns structured failure evidence.</summary>
    public static async Task<PdfCore.PdfSaveResult> SaveAsPdfResultAsync(this OdpPresentation source, Stream stream, PowerPointOpenDocumentConversionOptions? conversionOptions = null, PowerPointPdf.PowerPointToPdfOptions? pdfOptions = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try { return await source.ToPdfDocumentResult(conversionOptions, pdfOptions, cancellationToken).SaveResultAsync(stream, cancellationToken).ConfigureAwait(false); }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { throw; }
        catch (Exception ex) { return PdfCore.PdfSaveResult.FromFailure(outputPath: null, ex); }
    }

    /// <summary>Reconstructs an ODP presentation from an opened PDF.</summary>
    public static OdpPresentation ToOdpPresentation(this PdfCore.PdfDocument source, PowerPointPdf.PdfToPowerPointOptions? pdfOptions = null, PowerPointOpenDocumentConversionOptions? openDocumentOptions = null, System.Threading.CancellationToken cancellationToken = default) =>
        source.ToOdpPresentationResult(pdfOptions, openDocumentOptions, cancellationToken).Value;

    /// <summary>Reconstructs an ODP presentation and preserves diagnostics from both conversion stages.</summary>
    public static PdfOdpConversionResult ToOdpPresentationResult(this PdfCore.PdfDocument source, PowerPointPdf.PdfToPowerPointOptions? pdfOptions = null, PowerPointOpenDocumentConversionOptions? openDocumentOptions = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (source == null) throw new ArgumentNullException(nameof(source));
        PowerPointPdf.PdfPowerPointConversionResult pdf =
            PowerPointPdf.PowerPointPdfConverterExtensions.ToPowerPointPresentationResult(source, pdfOptions, cancellationToken);
        using (pdf.Value) {
            OdfConversionResult<OdpPresentation> odp = pdf.Value.ToOpenDocumentResult(openDocumentOptions);
            cancellationToken.ThrowIfCancellationRequested();
            return new PdfOdpConversionResult(
                odp.Value,
                new PdfOdpConversionReport(pdf.Report, odp.Report));
        }
    }

    /// <summary>Reconstructs an ODP presentation from an already loaded logical PDF model.</summary>
    /// <remarks>A logical PDF model resolves Auto to editable tables and also supports explicit editable-content reconstruction. Visual and hybrid projections require an opened PDF.</remarks>
    public static OdpPresentation ToOdpPresentation(this PdfCore.PdfDocumentReadResult source, PowerPointPdf.PdfToPowerPointOptions? pdfOptions = null, PowerPointOpenDocumentConversionOptions? openDocumentOptions = null, System.Threading.CancellationToken cancellationToken = default) =>
        source.ToOdpPresentationResult(pdfOptions, openDocumentOptions, cancellationToken).Value;

    /// <summary>Reconstructs an ODP presentation from a logical PDF model and preserves both stage reports.</summary>
    /// <remarks>A logical PDF model resolves Auto to editable tables and also supports explicit editable-content reconstruction. Visual and hybrid projections require an opened PDF.</remarks>
    public static PdfOdpConversionResult ToOdpPresentationResult(this PdfCore.PdfDocumentReadResult source, PowerPointPdf.PdfToPowerPointOptions? pdfOptions = null, PowerPointOpenDocumentConversionOptions? openDocumentOptions = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (source == null) throw new ArgumentNullException(nameof(source));
        PowerPointPdf.PdfPowerPointConversionResult pdf =
            PowerPointPdf.PowerPointPdfConverterExtensions.ToPowerPointPresentationResult(source, pdfOptions, cancellationToken);
        using (pdf.Value) {
            OdfConversionResult<OdpPresentation> odp = pdf.Value.ToOpenDocumentResult(openDocumentOptions);
            cancellationToken.ThrowIfCancellationRequested();
            return new PdfOdpConversionResult(
                odp.Value,
                new PdfOdpConversionReport(pdf.Report, odp.Report));
        }
    }

    /// <summary>Reconstructs and saves an ODP presentation from an opened PDF.</summary>
    public static OfficeOutputResult<PdfOdpConversionReport> SaveAsOdp(this PdfCore.PdfDocument source, string path, PowerPointPdf.PdfToPowerPointOptions? pdfOptions = null, PowerPointOpenDocumentConversionOptions? openDocumentOptions = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfOdpConversionResult result = source.ToOdpPresentationResult(pdfOptions, openDocumentOptions, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        result.Value.Save(path);
        return OfficeOutputResult<PdfOdpConversionReport>.FromSuccess(path, result.Report);
    }

    /// <summary>Reconstructs and writes an ODP presentation from an opened PDF.</summary>
    public static OfficeOutputResult<PdfOdpConversionReport> SaveAsOdp(this PdfCore.PdfDocument source, Stream stream, PowerPointPdf.PdfToPowerPointOptions? pdfOptions = null, PowerPointOpenDocumentConversionOptions? openDocumentOptions = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfOdpConversionResult result = source.ToOdpPresentationResult(pdfOptions, openDocumentOptions, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        result.Value.Save(stream);
        return OfficeOutputResult<PdfOdpConversionReport>.FromSuccess(null, result.Report);
    }

    /// <summary>Reconstructs and saves an ODP presentation from a logical PDF model.</summary>
    public static OfficeOutputResult<PdfOdpConversionReport> SaveAsOdp(this PdfCore.PdfDocumentReadResult source, string path, PowerPointPdf.PdfToPowerPointOptions? pdfOptions = null, PowerPointOpenDocumentConversionOptions? openDocumentOptions = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfOdpConversionResult result = source.ToOdpPresentationResult(pdfOptions, openDocumentOptions, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        result.Value.Save(path);
        return OfficeOutputResult<PdfOdpConversionReport>.FromSuccess(path, result.Report);
    }

    /// <summary>Reconstructs and writes an ODP presentation from a logical PDF model.</summary>
    public static OfficeOutputResult<PdfOdpConversionReport> SaveAsOdp(this PdfCore.PdfDocumentReadResult source, Stream stream, PowerPointPdf.PdfToPowerPointOptions? pdfOptions = null, PowerPointOpenDocumentConversionOptions? openDocumentOptions = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfOdpConversionResult result = source.ToOdpPresentationResult(pdfOptions, openDocumentOptions, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        result.Value.Save(stream);
        return OfficeOutputResult<PdfOdpConversionReport>.FromSuccess(null, result.Report);
    }

    /// <summary>Reconstructs synchronously, then asynchronously saves an ODP presentation from an opened PDF.</summary>
    public static async Task<OfficeOutputResult<PdfOdpConversionReport>> SaveAsOdpAsync(this PdfCore.PdfDocument source, string path, PowerPointPdf.PdfToPowerPointOptions? pdfOptions = null, PowerPointOpenDocumentConversionOptions? openDocumentOptions = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfOdpConversionResult result = source.ToOdpPresentationResult(pdfOptions, openDocumentOptions, cancellationToken);
        await result.Value.SaveAsync(path, cancellationToken: cancellationToken).ConfigureAwait(false);
        return OfficeOutputResult<PdfOdpConversionReport>.FromSuccess(path, result.Report);
    }

    /// <summary>Reconstructs synchronously, then asynchronously writes an ODP presentation from an opened PDF.</summary>
    public static async Task<OfficeOutputResult<PdfOdpConversionReport>> SaveAsOdpAsync(this PdfCore.PdfDocument source, Stream stream, PowerPointPdf.PdfToPowerPointOptions? pdfOptions = null, PowerPointOpenDocumentConversionOptions? openDocumentOptions = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfOdpConversionResult result = source.ToOdpPresentationResult(pdfOptions, openDocumentOptions, cancellationToken);
        await result.Value.SaveAsync(stream, cancellationToken: cancellationToken).ConfigureAwait(false);
        return OfficeOutputResult<PdfOdpConversionReport>.FromSuccess(null, result.Report);
    }

    /// <summary>Reconstructs synchronously, then asynchronously saves an ODP presentation from a logical PDF model.</summary>
    public static async Task<OfficeOutputResult<PdfOdpConversionReport>> SaveAsOdpAsync(this PdfCore.PdfDocumentReadResult source, string path, PowerPointPdf.PdfToPowerPointOptions? pdfOptions = null, PowerPointOpenDocumentConversionOptions? openDocumentOptions = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfOdpConversionResult result = source.ToOdpPresentationResult(pdfOptions, openDocumentOptions, cancellationToken);
        await result.Value.SaveAsync(path, cancellationToken: cancellationToken).ConfigureAwait(false);
        return OfficeOutputResult<PdfOdpConversionReport>.FromSuccess(path, result.Report);
    }

    /// <summary>Reconstructs synchronously, then asynchronously writes an ODP presentation from a logical PDF model.</summary>
    public static async Task<OfficeOutputResult<PdfOdpConversionReport>> SaveAsOdpAsync(this PdfCore.PdfDocumentReadResult source, Stream stream, PowerPointPdf.PdfToPowerPointOptions? pdfOptions = null, PowerPointOpenDocumentConversionOptions? openDocumentOptions = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfOdpConversionResult result = source.ToOdpPresentationResult(pdfOptions, openDocumentOptions, cancellationToken);
        await result.Value.SaveAsync(stream, cancellationToken: cancellationToken).ConfigureAwait(false);
        return OfficeOutputResult<PdfOdpConversionReport>.FromSuccess(null, result.Report);
    }
}
