using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Visio.Pdf;

/// <summary>Converts loaded Visio documents through the Visio-owned neutral projection and PDF engine.</summary>
public static class VisioPdfConverterExtensions {
    /// <summary>Converts a Visio document to the first-party PDF document model.</summary>
    public static PdfCore.PdfDocument ToPdfDocument(this VisioDocument document, VisioPdfSaveOptions? options = null) =>
        document.ToPdfDocumentResult(options).Value;

    /// <summary>Converts a Visio document and returns explicit conversion diagnostics.</summary>
    public static PdfCore.PdfDocumentConversionResult ToPdfDocumentResult(this VisioDocument document, VisioPdfSaveOptions? options = null) =>
        VisioPdfConversionEngine.Convert(document, options);

    /// <summary>Converts a Visio document to PDF bytes.</summary>
    public static byte[] ToPdf(this VisioDocument document, VisioPdfSaveOptions? options = null) =>
        document.ToPdfDocumentResult(options).ToBytes();

    /// <summary>Saves a Visio document as PDF and returns conversion diagnostics.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdf(this VisioDocument document, string path, VisioPdfSaveOptions? options = null) =>
        document.ToPdfDocumentResult(options).Save(path);

    /// <summary>Writes a Visio document as PDF to a caller-owned stream.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdf(this VisioDocument document, Stream stream, VisioPdfSaveOptions? options = null) =>
        document.ToPdfDocumentResult(options).Save(stream);

    /// <summary>Attempts to save a Visio document as PDF and returns structured failure evidence.</summary>
    public static PdfCore.PdfSaveResult TrySaveAsPdf(this VisioDocument document, string path, VisioPdfSaveOptions? options = null) {
        try { return document.ToPdfDocumentResult(options).TrySave(path); }
        catch (Exception exception) { return PdfCore.PdfSaveResult.FromFailure(path, exception); }
    }

    /// <summary>Attempts to write a Visio document as PDF to a caller-owned stream.</summary>
    public static PdfCore.PdfSaveResult TrySaveAsPdf(this VisioDocument document, Stream stream, VisioPdfSaveOptions? options = null) {
        try { return document.ToPdfDocumentResult(options).TrySave(stream); }
        catch (Exception exception) { return PdfCore.PdfSaveResult.FromFailure(null, exception); }
    }

    /// <summary>Converts synchronously, then asynchronously saves a Visio PDF to a path.</summary>
    public static Task<PdfCore.PdfSaveResult> SaveAsPdfAsync(
        this VisioDocument document,
        string path,
        VisioPdfSaveOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        return VisioPdfConversionEngine.Convert(document, options, cancellationToken).SaveAsync(path, cancellationToken);
    }

    /// <summary>Converts synchronously, then asynchronously writes a Visio PDF to a caller-owned stream.</summary>
    public static Task<PdfCore.PdfSaveResult> SaveAsPdfAsync(
        this VisioDocument document,
        Stream stream,
        VisioPdfSaveOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        return VisioPdfConversionEngine.Convert(document, options, cancellationToken).SaveAsync(stream, cancellationToken);
    }

    /// <summary>Attempts to asynchronously save a Visio PDF and returns structured failure evidence.</summary>
    public static async Task<PdfCore.PdfSaveResult> TrySaveAsPdfAsync(
        this VisioDocument document,
        string path,
        VisioPdfSaveOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try {
            return await VisioPdfConversionEngine.Convert(document, options, cancellationToken)
                .TrySaveAsync(path, cancellationToken)
                .ConfigureAwait(false);
        } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
            throw;
        } catch (Exception exception) {
            return PdfCore.PdfSaveResult.FromFailure(path, exception);
        }
    }

    /// <summary>Attempts to asynchronously write a Visio PDF and returns structured failure evidence.</summary>
    public static async Task<PdfCore.PdfSaveResult> TrySaveAsPdfAsync(
        this VisioDocument document,
        Stream stream,
        VisioPdfSaveOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try {
            return await VisioPdfConversionEngine.Convert(document, options, cancellationToken)
                .TrySaveAsync(stream, cancellationToken)
                .ConfigureAwait(false);
        } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
            throw;
        } catch (Exception exception) {
            return PdfCore.PdfSaveResult.FromFailure(null, exception);
        }
    }
}
