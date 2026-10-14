using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.OpenDocument.Odg.Pdf;

/// <summary>Converts loaded Draw documents through the Draw-owned Drawing projection and PDF engine.</summary>
public static class OdgPdfConversionExtensions {
    /// <summary>Converts a Draw document to the first-party PDF document model.</summary>
    public static PdfCore.PdfDocument ToPdfDocument(this OdgDocument document, OdgToPdfOptions? options = null, System.Threading.CancellationToken cancellationToken = default) =>
        document.ToPdfDocumentResult(options, cancellationToken).Value;

    /// <summary>Converts a Draw document and returns explicit conversion diagnostics.</summary>
    public static PdfCore.PdfDocumentConversionResult ToPdfDocumentResult(this OdgDocument document, OdgToPdfOptions? options = null, System.Threading.CancellationToken cancellationToken = default) =>
        OdgPdfConversionEngine.Convert(document, options, cancellationToken);

    /// <summary>Converts a Draw document to PDF bytes.</summary>
    public static byte[] ToPdfBytes(this OdgDocument document, OdgToPdfOptions? options = null, System.Threading.CancellationToken cancellationToken = default) =>
        document.ToPdfDocumentResult(options, cancellationToken).ToBytes(cancellationToken);

    /// <summary>Saves a Draw document as PDF and returns conversion diagnostics.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdf(this OdgDocument document, string path, OdgToPdfOptions? options = null, System.Threading.CancellationToken cancellationToken = default) =>
        document.ToPdfDocumentResult(options, cancellationToken).Save(path, cancellationToken);

    /// <summary>Writes a Draw document as PDF to a caller-owned stream.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdf(this OdgDocument document, Stream stream, OdgToPdfOptions? options = null, System.Threading.CancellationToken cancellationToken = default) =>
        document.ToPdfDocumentResult(options, cancellationToken).Save(stream, cancellationToken);

    /// <summary>Attempts to save a Draw document as PDF and returns structured failure evidence.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdfResult(this OdgDocument document, string path, OdgToPdfOptions? options = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try { return document.ToPdfDocumentResult(options, cancellationToken).SaveResult(path, cancellationToken); }
        catch (OperationCanceledException) { throw; }
        catch (OdfConversionLossException exception) { return PdfCore.PdfSaveResult.FromFailure(path, exception, exception.Report); }
        catch (Exception exception) { return PdfCore.PdfSaveResult.FromFailure(path, exception); }
    }

    /// <summary>Attempts to write a Draw document as PDF to a caller-owned stream.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdfResult(this OdgDocument document, Stream stream, OdgToPdfOptions? options = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try { return document.ToPdfDocumentResult(options, cancellationToken).SaveResult(stream, cancellationToken); }
        catch (OperationCanceledException) { throw; }
        catch (OdfConversionLossException exception) { return PdfCore.PdfSaveResult.FromFailure(null, exception, exception.Report); }
        catch (Exception exception) { return PdfCore.PdfSaveResult.FromFailure(null, exception); }
    }

    /// <summary>Converts synchronously, then asynchronously saves a Draw PDF to a path.</summary>
    public static Task<PdfCore.PdfSaveResult> SaveAsPdfAsync(
        this OdgDocument document,
        string path,
        OdgToPdfOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        return OdgPdfConversionEngine.Convert(document, options, cancellationToken).SaveAsync(path, cancellationToken);
    }

    /// <summary>Converts synchronously, then asynchronously writes a Draw PDF to a caller-owned stream.</summary>
    public static Task<PdfCore.PdfSaveResult> SaveAsPdfAsync(
        this OdgDocument document,
        Stream stream,
        OdgToPdfOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        return OdgPdfConversionEngine.Convert(document, options, cancellationToken).SaveAsync(stream, cancellationToken);
    }

    /// <summary>Attempts to asynchronously save a Draw PDF and returns structured failure evidence.</summary>
    public static async Task<PdfCore.PdfSaveResult> SaveAsPdfResultAsync(
        this OdgDocument document,
        string path,
        OdgToPdfOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try {
            return await OdgPdfConversionEngine.Convert(document, options, cancellationToken)
                .SaveResultAsync(path, cancellationToken)
                .ConfigureAwait(false);
        } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
            throw;
        } catch (OdfConversionLossException exception) {
            return PdfCore.PdfSaveResult.FromFailure(path, exception, exception.Report);
        } catch (Exception exception) {
            return PdfCore.PdfSaveResult.FromFailure(path, exception);
        }
    }

    /// <summary>Attempts to asynchronously write a Draw PDF and returns structured failure evidence.</summary>
    public static async Task<PdfCore.PdfSaveResult> SaveAsPdfResultAsync(
        this OdgDocument document,
        Stream stream,
        OdgToPdfOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try {
            return await OdgPdfConversionEngine.Convert(document, options, cancellationToken)
                .SaveResultAsync(stream, cancellationToken)
                .ConfigureAwait(false);
        } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
            throw;
        } catch (OdfConversionLossException exception) {
            return PdfCore.PdfSaveResult.FromFailure(null, exception, exception.Report);
        } catch (Exception exception) {
            return PdfCore.PdfSaveResult.FromFailure(null, exception);
        }
    }
}
