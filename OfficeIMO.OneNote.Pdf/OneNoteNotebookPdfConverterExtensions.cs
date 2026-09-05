using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.OneNote.Pdf;

/// <summary>Converts offline OneNote notebooks to semantic PDF documents.</summary>
public static class OneNoteNotebookPdfConverterExtensions {
    /// <summary>Converts a notebook to a first-party PDF document.</summary>
    public static PdfCore.PdfDocument ToPdfDocument(this OneNoteNotebook notebook, OneNoteToPdfOptions? options = null, System.Threading.CancellationToken cancellationToken = default) =>
        notebook.ToPdfDocumentResult(options, cancellationToken).Value;

    /// <summary>Converts a notebook and returns explicit source, projection, and PDF diagnostics.</summary>
    public static PdfCore.PdfDocumentConversionResult ToPdfDocumentResult(this OneNoteNotebook notebook, OneNoteToPdfOptions? options = null, System.Threading.CancellationToken cancellationToken = default) =>
        OneNotePdfConversionEngine.Convert(notebook, options, cancellationToken);

    /// <summary>Converts a notebook to PDF bytes.</summary>
    public static byte[] ToPdfBytes(this OneNoteNotebook notebook, OneNoteToPdfOptions? options = null, System.Threading.CancellationToken cancellationToken = default) =>
        notebook.ToPdfDocumentResult(options, cancellationToken).ToBytes(cancellationToken);

    /// <summary>Saves a notebook as PDF and returns conversion diagnostics.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdf(this OneNoteNotebook notebook, string path, OneNoteToPdfOptions? options = null, System.Threading.CancellationToken cancellationToken = default) =>
        notebook.ToPdfDocumentResult(options, cancellationToken).Save(path, cancellationToken);

    /// <summary>Writes a notebook as PDF to a caller-owned stream.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdf(this OneNoteNotebook notebook, Stream stream, OneNoteToPdfOptions? options = null, System.Threading.CancellationToken cancellationToken = default) =>
        notebook.ToPdfDocumentResult(options, cancellationToken).Save(stream, cancellationToken);

    /// <summary>Attempts to save a notebook as PDF and returns structured failure evidence.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdfResult(this OneNoteNotebook notebook, string path, OneNoteToPdfOptions? options = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try { return notebook.ToPdfDocumentResult(options, cancellationToken).SaveResult(path, cancellationToken); }
        catch (OperationCanceledException) { throw; }
        catch (Exception exception) { return PdfCore.PdfSaveResult.FromFailure(path, exception); }
    }

    /// <summary>Attempts to write a notebook as PDF to a caller-owned stream.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdfResult(this OneNoteNotebook notebook, Stream stream, OneNoteToPdfOptions? options = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try { return notebook.ToPdfDocumentResult(options, cancellationToken).SaveResult(stream, cancellationToken); }
        catch (OperationCanceledException) { throw; }
        catch (Exception exception) { return PdfCore.PdfSaveResult.FromFailure(null, exception); }
    }

    /// <summary>Converts synchronously, then asynchronously writes a notebook PDF to a path.</summary>
    public static Task<PdfCore.PdfSaveResult> SaveAsPdfAsync(this OneNoteNotebook notebook, string path, OneNoteToPdfOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        return notebook.ToPdfDocumentResult(options, cancellationToken).SaveAsync(path, cancellationToken);
    }

    /// <summary>Converts synchronously, then asynchronously writes a notebook PDF to a caller-owned stream.</summary>
    public static Task<PdfCore.PdfSaveResult> SaveAsPdfAsync(this OneNoteNotebook notebook, Stream stream, OneNoteToPdfOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        return notebook.ToPdfDocumentResult(options, cancellationToken).SaveAsync(stream, cancellationToken);
    }

    /// <summary>Converts synchronously, then attempts to asynchronously save a notebook PDF.</summary>
    public static async Task<PdfCore.PdfSaveResult> SaveAsPdfResultAsync(this OneNoteNotebook notebook, string path, OneNoteToPdfOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try { return await notebook.ToPdfDocumentResult(options, cancellationToken).SaveResultAsync(path, cancellationToken).ConfigureAwait(false); }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { throw; }
        catch (Exception exception) { return PdfCore.PdfSaveResult.FromFailure(path, exception); }
    }

    /// <summary>Converts synchronously, then attempts to asynchronously write a notebook PDF.</summary>
    public static async Task<PdfCore.PdfSaveResult> SaveAsPdfResultAsync(this OneNoteNotebook notebook, Stream stream, OneNoteToPdfOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try { return await notebook.ToPdfDocumentResult(options, cancellationToken).SaveResultAsync(stream, cancellationToken).ConfigureAwait(false); }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { throw; }
        catch (Exception exception) { return PdfCore.PdfSaveResult.FromFailure(null, exception); }
    }
}
