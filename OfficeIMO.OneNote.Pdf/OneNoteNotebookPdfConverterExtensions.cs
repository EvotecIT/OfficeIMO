using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.OneNote.Pdf;

/// <summary>Converts offline OneNote notebooks to semantic PDF documents.</summary>
public static class OneNoteNotebookPdfConverterExtensions {
    /// <summary>Converts a notebook to a first-party PDF document.</summary>
    public static PdfCore.PdfDocument ToPdfDocument(this OneNoteNotebook notebook, OneNotePdfSaveOptions? options = null) =>
        notebook.ToPdfDocumentResult(options).Value;

    /// <summary>Converts a notebook and returns explicit source, projection, and PDF diagnostics.</summary>
    public static PdfCore.PdfDocumentConversionResult ToPdfDocumentResult(this OneNoteNotebook notebook, OneNotePdfSaveOptions? options = null) =>
        OneNotePdfConversionEngine.Convert(notebook, options);

    /// <summary>Converts a notebook to PDF bytes.</summary>
    public static byte[] ToPdf(this OneNoteNotebook notebook, OneNotePdfSaveOptions? options = null) =>
        notebook.ToPdfDocumentResult(options).ToBytes();

    /// <summary>Saves a notebook as PDF and returns conversion diagnostics.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdf(this OneNoteNotebook notebook, string path, OneNotePdfSaveOptions? options = null) =>
        notebook.ToPdfDocumentResult(options).Save(path);

    /// <summary>Writes a notebook as PDF to a caller-owned stream.</summary>
    public static PdfCore.PdfSaveResult SaveAsPdf(this OneNoteNotebook notebook, Stream stream, OneNotePdfSaveOptions? options = null) =>
        notebook.ToPdfDocumentResult(options).Save(stream);

    /// <summary>Attempts to save a notebook as PDF and returns structured failure evidence.</summary>
    public static PdfCore.PdfSaveResult TrySaveAsPdf(this OneNoteNotebook notebook, string path, OneNotePdfSaveOptions? options = null) {
        try { return notebook.ToPdfDocumentResult(options).TrySave(path); }
        catch (Exception exception) { return PdfCore.PdfSaveResult.FromFailure(path, exception); }
    }

    /// <summary>Attempts to write a notebook as PDF to a caller-owned stream.</summary>
    public static PdfCore.PdfSaveResult TrySaveAsPdf(this OneNoteNotebook notebook, Stream stream, OneNotePdfSaveOptions? options = null) {
        try { return notebook.ToPdfDocumentResult(options).TrySave(stream); }
        catch (Exception exception) { return PdfCore.PdfSaveResult.FromFailure(null, exception); }
    }

    /// <summary>Converts synchronously, then asynchronously writes a notebook PDF to a path.</summary>
    public static Task<PdfCore.PdfSaveResult> SaveAsPdfAsync(this OneNoteNotebook notebook, string path, OneNotePdfSaveOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        return notebook.ToPdfDocumentResult(options).SaveAsync(path, cancellationToken);
    }

    /// <summary>Converts synchronously, then asynchronously writes a notebook PDF to a caller-owned stream.</summary>
    public static Task<PdfCore.PdfSaveResult> SaveAsPdfAsync(this OneNoteNotebook notebook, Stream stream, OneNotePdfSaveOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        return notebook.ToPdfDocumentResult(options).SaveAsync(stream, cancellationToken);
    }

    /// <summary>Converts synchronously, then attempts to asynchronously save a notebook PDF.</summary>
    public static async Task<PdfCore.PdfSaveResult> TrySaveAsPdfAsync(this OneNoteNotebook notebook, string path, OneNotePdfSaveOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try { return await notebook.ToPdfDocumentResult(options).TrySaveAsync(path, cancellationToken).ConfigureAwait(false); }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { throw; }
        catch (Exception exception) { return PdfCore.PdfSaveResult.FromFailure(path, exception); }
    }

    /// <summary>Converts synchronously, then attempts to asynchronously write a notebook PDF.</summary>
    public static async Task<PdfCore.PdfSaveResult> TrySaveAsPdfAsync(this OneNoteNotebook notebook, Stream stream, OneNotePdfSaveOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try { return await notebook.ToPdfDocumentResult(options).TrySaveAsync(stream, cancellationToken).ConfigureAwait(false); }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { throw; }
        catch (Exception exception) { return PdfCore.PdfSaveResult.FromFailure(null, exception); }
    }
}
