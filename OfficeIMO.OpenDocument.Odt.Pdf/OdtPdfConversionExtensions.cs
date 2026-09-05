using OfficeIMO.Word.OpenDocument;
using WordPdf = OfficeIMO.Word.Pdf;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.OpenDocument.Odt.Pdf;

/// <summary>Direct, loss-aware ODT to PDF conversion through the Word semantic and PDF engines.</summary>
public static class OdtPdfConversionExtensions {
    /// <summary>Converts an ODT document to the first-party PDF document model.</summary>
    public static PdfCore.PdfDocument ToPdfDocument(
        this OdtDocument source,
        WordOpenDocumentConversionOptions? conversionOptions = null,
        WordPdf.WordToPdfOptions? pdfOptions = null, System.Threading.CancellationToken cancellationToken = default) =>
        source.ToPdfDocumentResult(conversionOptions, pdfOptions, cancellationToken).Value;

    /// <summary>Converts an ODT document to PDF and preserves diagnostics from both conversion stages.</summary>
    public static PdfCore.PdfDocumentConversionResult ToPdfDocumentResult(
        this OdtDocument source,
        WordOpenDocumentConversionOptions? conversionOptions = null,
        WordPdf.WordToPdfOptions? pdfOptions = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (source == null) throw new ArgumentNullException(nameof(source));
        OdfConversionResult<OfficeIMO.Word.WordDocument> conversion =
            source.ToWordDocumentResult(conversionOptions);
        using (conversion.Value) {
            PdfCore.PdfDocumentConversionResult result =
                WordPdf.WordPdfConverterExtensions.ToPdfDocumentResult(conversion.Value, pdfOptions, cancellationToken);
            return result.WithSourceConversionReport(conversion.Report);
        }
    }

    /// <summary>Converts an ODT document to PDF bytes.</summary>
    public static byte[] ToPdfBytes(this OdtDocument source, WordOpenDocumentConversionOptions? conversionOptions = null, WordPdf.WordToPdfOptions? pdfOptions = null, System.Threading.CancellationToken cancellationToken = default) =>
        source.ToPdfDocumentResult(conversionOptions, pdfOptions, cancellationToken).ToBytes(cancellationToken);

    /// <summary>Saves an ODT document as PDF.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdf(this OdtDocument source, string path, WordOpenDocumentConversionOptions? conversionOptions = null, WordPdf.WordToPdfOptions? pdfOptions = null, System.Threading.CancellationToken cancellationToken = default) =>
        source.ToPdfDocumentResult(conversionOptions, pdfOptions, cancellationToken).Save(path, cancellationToken);

    /// <summary>Writes an ODT document as PDF to a caller-owned stream.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdf(this OdtDocument source, Stream stream, WordOpenDocumentConversionOptions? conversionOptions = null, WordPdf.WordToPdfOptions? pdfOptions = null, System.Threading.CancellationToken cancellationToken = default) =>
        source.ToPdfDocumentResult(conversionOptions, pdfOptions, cancellationToken).Save(stream, cancellationToken);

    /// <summary>Attempts to save an ODT document as PDF and returns structured failure evidence.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdfResult(this OdtDocument source, string path, WordOpenDocumentConversionOptions? conversionOptions = null, WordPdf.WordToPdfOptions? pdfOptions = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try { return source.ToPdfDocumentResult(conversionOptions, pdfOptions, cancellationToken).SaveResult(path, cancellationToken); }
        catch (OperationCanceledException) { throw; }
        catch (Exception ex) { return PdfCore.PdfSaveResult.FromFailure(path, ex); }
    }

    /// <summary>Attempts to write an ODT document as PDF and returns structured failure evidence.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdfResult(this OdtDocument source, Stream stream, WordOpenDocumentConversionOptions? conversionOptions = null, WordPdf.WordToPdfOptions? pdfOptions = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try { return source.ToPdfDocumentResult(conversionOptions, pdfOptions, cancellationToken).SaveResult(stream, cancellationToken); }
        catch (OperationCanceledException) { throw; }
        catch (Exception ex) { return PdfCore.PdfSaveResult.FromFailure(outputPath: null, ex); }
    }

    /// <summary>Converts synchronously, then asynchronously saves an ODT document as PDF.</summary>
    public static Task<PdfCore.PdfSaveResult> SaveAsPdfAsync(this OdtDocument source, string path, WordOpenDocumentConversionOptions? conversionOptions = null, WordPdf.WordToPdfOptions? pdfOptions = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        return source.ToPdfDocumentResult(conversionOptions, pdfOptions, cancellationToken).SaveAsync(path, cancellationToken);
    }

    /// <summary>Converts synchronously, then asynchronously writes an ODT document as PDF.</summary>
    public static Task<PdfCore.PdfSaveResult> SaveAsPdfAsync(this OdtDocument source, Stream stream, WordOpenDocumentConversionOptions? conversionOptions = null, WordPdf.WordToPdfOptions? pdfOptions = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        return source.ToPdfDocumentResult(conversionOptions, pdfOptions, cancellationToken).SaveAsync(stream, cancellationToken);
    }

    /// <summary>Attempts to asynchronously save an ODT document as PDF and returns structured failure evidence.</summary>
    public static async Task<PdfCore.PdfSaveResult> SaveAsPdfResultAsync(this OdtDocument source, string path, WordOpenDocumentConversionOptions? conversionOptions = null, WordPdf.WordToPdfOptions? pdfOptions = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try { return await source.ToPdfDocumentResult(conversionOptions, pdfOptions, cancellationToken).SaveResultAsync(path, cancellationToken).ConfigureAwait(false); }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { throw; }
        catch (Exception ex) { return PdfCore.PdfSaveResult.FromFailure(path, ex); }
    }

    /// <summary>Attempts to asynchronously write an ODT document as PDF and returns structured failure evidence.</summary>
    public static async Task<PdfCore.PdfSaveResult> SaveAsPdfResultAsync(this OdtDocument source, Stream stream, WordOpenDocumentConversionOptions? conversionOptions = null, WordPdf.WordToPdfOptions? pdfOptions = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try { return await source.ToPdfDocumentResult(conversionOptions, pdfOptions, cancellationToken).SaveResultAsync(stream, cancellationToken).ConfigureAwait(false); }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { throw; }
        catch (Exception ex) { return PdfCore.PdfSaveResult.FromFailure(outputPath: null, ex); }
    }

    /// <summary>Reconstructs an ODT document from an opened PDF.</summary>
    public static OdtDocument ToOdtDocument(
        this PdfCore.PdfDocument source,
        WordPdf.PdfToWordOptions? pdfOptions = null,
        WordOpenDocumentConversionOptions? openDocumentOptions = null, System.Threading.CancellationToken cancellationToken = default) =>
        source.ToOdtDocumentResult(pdfOptions, openDocumentOptions, cancellationToken).Value;

    /// <summary>Reconstructs an ODT document and preserves diagnostics from both semantic stages.</summary>
    public static PdfOdtConversionResult ToOdtDocumentResult(
        this PdfCore.PdfDocument source,
        WordPdf.PdfToWordOptions? pdfOptions = null,
        WordOpenDocumentConversionOptions? openDocumentOptions = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (source == null) throw new ArgumentNullException(nameof(source));
        WordPdf.PdfWordConversionResult pdf =
            WordPdf.PdfWordConverterExtensions.ToWordDocumentResult(source, pdfOptions, cancellationToken);
        using (pdf.Value) {
            OdfConversionResult<OdtDocument> odt = pdf.Value.ToOpenDocumentResult(openDocumentOptions);
            return new PdfOdtConversionResult(
                odt.Value,
                new PdfOdtConversionReport(pdf.Report, odt.Report));
        }
    }

    /// <summary>Reconstructs an ODT document from an already loaded logical PDF model.</summary>
    public static OdtDocument ToOdtDocument(
        this PdfCore.PdfDocumentReadResult source,
        WordPdf.PdfToWordOptions? pdfOptions = null,
        WordOpenDocumentConversionOptions? openDocumentOptions = null, System.Threading.CancellationToken cancellationToken = default) =>
        source.ToOdtDocumentResult(pdfOptions, openDocumentOptions, cancellationToken).Value;

    /// <summary>Reconstructs an ODT document from a logical PDF model and preserves both stage reports.</summary>
    public static PdfOdtConversionResult ToOdtDocumentResult(
        this PdfCore.PdfDocumentReadResult source,
        WordPdf.PdfToWordOptions? pdfOptions = null,
        WordOpenDocumentConversionOptions? openDocumentOptions = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (source == null) throw new ArgumentNullException(nameof(source));
        WordPdf.PdfWordConversionResult pdf =
            WordPdf.PdfWordConverterExtensions.ToWordDocumentResult(source, pdfOptions, cancellationToken);
        using (pdf.Value) {
            OdfConversionResult<OdtDocument> odt = pdf.Value.ToOpenDocumentResult(openDocumentOptions);
            return new PdfOdtConversionResult(
                odt.Value,
                new PdfOdtConversionReport(pdf.Report, odt.Report));
        }
    }

    /// <summary>Reconstructs and saves an ODT document from an opened PDF.</summary>
    public static OfficeOutputResult<PdfOdtConversionReport> SaveAsOdt(this PdfCore.PdfDocument source, string path, WordPdf.PdfToWordOptions? pdfOptions = null, WordOpenDocumentConversionOptions? openDocumentOptions = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfOdtConversionResult result = source.ToOdtDocumentResult(pdfOptions, openDocumentOptions, cancellationToken);
        result.Value.Save(path);
        return OfficeOutputResult<PdfOdtConversionReport>.FromSuccess(path, result.Report);
    }

    /// <summary>Reconstructs and writes an ODT document from an opened PDF.</summary>
    public static OfficeOutputResult<PdfOdtConversionReport> SaveAsOdt(this PdfCore.PdfDocument source, Stream stream, WordPdf.PdfToWordOptions? pdfOptions = null, WordOpenDocumentConversionOptions? openDocumentOptions = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfOdtConversionResult result = source.ToOdtDocumentResult(pdfOptions, openDocumentOptions, cancellationToken);
        result.Value.Save(stream);
        return OfficeOutputResult<PdfOdtConversionReport>.FromSuccess(null, result.Report);
    }

    /// <summary>Reconstructs and saves an ODT document from a logical PDF model.</summary>
    public static OfficeOutputResult<PdfOdtConversionReport> SaveAsOdt(this PdfCore.PdfDocumentReadResult source, string path, WordPdf.PdfToWordOptions? pdfOptions = null, WordOpenDocumentConversionOptions? openDocumentOptions = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfOdtConversionResult result = source.ToOdtDocumentResult(pdfOptions, openDocumentOptions, cancellationToken);
        result.Value.Save(path);
        return OfficeOutputResult<PdfOdtConversionReport>.FromSuccess(path, result.Report);
    }

    /// <summary>Reconstructs and writes an ODT document from a logical PDF model.</summary>
    public static OfficeOutputResult<PdfOdtConversionReport> SaveAsOdt(this PdfCore.PdfDocumentReadResult source, Stream stream, WordPdf.PdfToWordOptions? pdfOptions = null, WordOpenDocumentConversionOptions? openDocumentOptions = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfOdtConversionResult result = source.ToOdtDocumentResult(pdfOptions, openDocumentOptions, cancellationToken);
        result.Value.Save(stream);
        return OfficeOutputResult<PdfOdtConversionReport>.FromSuccess(null, result.Report);
    }

    /// <summary>Reconstructs synchronously, then asynchronously saves an ODT document from an opened PDF.</summary>
    public static async Task<OfficeOutputResult<PdfOdtConversionReport>> SaveAsOdtAsync(this PdfCore.PdfDocument source, string path, WordPdf.PdfToWordOptions? pdfOptions = null, WordOpenDocumentConversionOptions? openDocumentOptions = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfOdtConversionResult result = source.ToOdtDocumentResult(pdfOptions, openDocumentOptions, cancellationToken);
        await result.Value.SaveAsync(path, cancellationToken: cancellationToken).ConfigureAwait(false);
        return OfficeOutputResult<PdfOdtConversionReport>.FromSuccess(path, result.Report);
    }

    /// <summary>Reconstructs synchronously, then asynchronously writes an ODT document from an opened PDF.</summary>
    public static async Task<OfficeOutputResult<PdfOdtConversionReport>> SaveAsOdtAsync(this PdfCore.PdfDocument source, Stream stream, WordPdf.PdfToWordOptions? pdfOptions = null, WordOpenDocumentConversionOptions? openDocumentOptions = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfOdtConversionResult result = source.ToOdtDocumentResult(pdfOptions, openDocumentOptions, cancellationToken);
        await result.Value.SaveAsync(stream, cancellationToken: cancellationToken).ConfigureAwait(false);
        return OfficeOutputResult<PdfOdtConversionReport>.FromSuccess(null, result.Report);
    }

    /// <summary>Reconstructs synchronously, then asynchronously saves an ODT document from a logical PDF model.</summary>
    public static async Task<OfficeOutputResult<PdfOdtConversionReport>> SaveAsOdtAsync(this PdfCore.PdfDocumentReadResult source, string path, WordPdf.PdfToWordOptions? pdfOptions = null, WordOpenDocumentConversionOptions? openDocumentOptions = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfOdtConversionResult result = source.ToOdtDocumentResult(pdfOptions, openDocumentOptions, cancellationToken);
        await result.Value.SaveAsync(path, cancellationToken: cancellationToken).ConfigureAwait(false);
        return OfficeOutputResult<PdfOdtConversionReport>.FromSuccess(path, result.Report);
    }

    /// <summary>Reconstructs synchronously, then asynchronously writes an ODT document from a logical PDF model.</summary>
    public static async Task<OfficeOutputResult<PdfOdtConversionReport>> SaveAsOdtAsync(this PdfCore.PdfDocumentReadResult source, Stream stream, WordPdf.PdfToWordOptions? pdfOptions = null, WordOpenDocumentConversionOptions? openDocumentOptions = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfOdtConversionResult result = source.ToOdtDocumentResult(pdfOptions, openDocumentOptions, cancellationToken);
        await result.Value.SaveAsync(stream, cancellationToken: cancellationToken).ConfigureAwait(false);
        return OfficeOutputResult<PdfOdtConversionReport>.FromSuccess(null, result.Report);
    }
}
