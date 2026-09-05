using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Visio.Pdf;

/// <summary>Converts loaded Visio documents through the Visio-owned neutral projection and PDF engine.</summary>
public static class VisioPdfConverterExtensions {
    /// <summary>Converts a Visio document to the first-party PDF document model.</summary>
    public static PdfCore.PdfDocument ToPdfDocument(this VisioDocument document, VisioToPdfOptions? options = null, System.Threading.CancellationToken cancellationToken = default) =>
        document.ToPdfDocumentResult(options, cancellationToken).Value;

    /// <summary>Converts a Visio document and returns explicit conversion diagnostics.</summary>
    public static PdfCore.PdfDocumentConversionResult ToPdfDocumentResult(this VisioDocument document, VisioToPdfOptions? options = null, System.Threading.CancellationToken cancellationToken = default) =>
        VisioPdfConversionEngine.Convert(document, options, cancellationToken);

    /// <summary>Converts a Visio document to PDF bytes.</summary>
    public static byte[] ToPdfBytes(this VisioDocument document, VisioToPdfOptions? options = null, System.Threading.CancellationToken cancellationToken = default) =>
        document.ToPdfDocumentResult(options, cancellationToken).ToBytes(cancellationToken);

    /// <summary>Saves a Visio document as PDF and returns conversion diagnostics.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdf(this VisioDocument document, string path, VisioToPdfOptions? options = null, System.Threading.CancellationToken cancellationToken = default) =>
        document.ToPdfDocumentResult(options, cancellationToken).Save(path, cancellationToken);

    /// <summary>Writes a Visio document as PDF to a caller-owned stream.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdf(this VisioDocument document, Stream stream, VisioToPdfOptions? options = null, System.Threading.CancellationToken cancellationToken = default) =>
        document.ToPdfDocumentResult(options, cancellationToken).Save(stream, cancellationToken);

    /// <summary>Attempts to save a Visio document as PDF and returns structured failure evidence.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdfResult(this VisioDocument document, string path, VisioToPdfOptions? options = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try { return document.ToPdfDocumentResult(options, cancellationToken).SaveResult(path, cancellationToken); }
        catch (OperationCanceledException) { throw; }
        catch (Exception exception) { return PdfCore.PdfSaveResult.FromFailure(path, exception); }
    }

    /// <summary>Attempts to write a Visio document as PDF to a caller-owned stream.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdfResult(this VisioDocument document, Stream stream, VisioToPdfOptions? options = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try { return document.ToPdfDocumentResult(options, cancellationToken).SaveResult(stream, cancellationToken); }
        catch (OperationCanceledException) { throw; }
        catch (Exception exception) { return PdfCore.PdfSaveResult.FromFailure(null, exception); }
    }

    /// <summary>Converts synchronously, then asynchronously saves a Visio PDF to a path.</summary>
    public static Task<PdfCore.PdfSaveResult> SaveAsPdfAsync(
        this VisioDocument document,
        string path,
        VisioToPdfOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        return VisioPdfConversionEngine.Convert(document, options, cancellationToken).SaveAsync(path, cancellationToken);
    }

    /// <summary>Converts synchronously, then asynchronously writes a Visio PDF to a caller-owned stream.</summary>
    public static Task<PdfCore.PdfSaveResult> SaveAsPdfAsync(
        this VisioDocument document,
        Stream stream,
        VisioToPdfOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        return VisioPdfConversionEngine.Convert(document, options, cancellationToken).SaveAsync(stream, cancellationToken);
    }

    /// <summary>Attempts to asynchronously save a Visio PDF and returns structured failure evidence.</summary>
    public static async Task<PdfCore.PdfSaveResult> SaveAsPdfResultAsync(
        this VisioDocument document,
        string path,
        VisioToPdfOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try {
            return await VisioPdfConversionEngine.Convert(document, options, cancellationToken)
                .SaveResultAsync(path, cancellationToken)
                .ConfigureAwait(false);
        } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
            throw;
        } catch (Exception exception) {
            return PdfCore.PdfSaveResult.FromFailure(path, exception);
        }
    }

    /// <summary>Attempts to asynchronously write a Visio PDF and returns structured failure evidence.</summary>
    public static async Task<PdfCore.PdfSaveResult> SaveAsPdfResultAsync(
        this VisioDocument document,
        Stream stream,
        VisioToPdfOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try {
            return await VisioPdfConversionEngine.Convert(document, options, cancellationToken)
                .SaveResultAsync(stream, cancellationToken)
                .ConfigureAwait(false);
        } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
            throw;
        } catch (Exception exception) {
            return PdfCore.PdfSaveResult.FromFailure(null, exception);
        }
    }
}
