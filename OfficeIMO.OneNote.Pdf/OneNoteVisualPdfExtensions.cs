using System.Collections.Generic;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.OneNote.Pdf;

/// <summary>Visual-preserving PDF entry points backed by the native OneNote Drawing canvas.</summary>
public static class OneNoteVisualPdfExtensions {
    /// <summary>Converts a section to a visual PDF document with rendering diagnostics.</summary>
    public static PdfCore.PdfDocumentConversionResult ToVisualPdfDocumentResult(this OneNoteSection section, OneNoteVisualPdfOptions? options = null, CancellationToken cancellationToken = default) =>
        RenderVisualPdf(section, options, cancellationToken);

    /// <summary>Converts a section to a visual PDF document.</summary>
    public static PdfCore.PdfDocument ToVisualPdfDocument(this OneNoteSection section, OneNoteVisualPdfOptions? options = null, CancellationToken cancellationToken = default) =>
        section.ToVisualPdfDocumentResult(options, cancellationToken).Value;

    /// <summary>Converts a section to serialized visual PDF bytes.</summary>
    public static byte[] ToVisualPdfBytes(this OneNoteSection section, OneNoteVisualPdfOptions? options = null, CancellationToken cancellationToken = default) =>
        section.ToVisualPdfDocumentResult(options, cancellationToken).ToBytes(cancellationToken);

    /// <summary>Writes a section as visual PDF and throws on failure.</summary>
    public static PdfCore.PdfSaveResult SaveAsVisualPdf(this OneNoteSection section, string path, OneNoteVisualPdfOptions? options = null, CancellationToken cancellationToken = default) =>
        section.ToVisualPdfDocumentResult(options, cancellationToken).Save(path, cancellationToken);

    /// <summary>Writes a section as visual PDF and captures failure evidence; cancellation still throws.</summary>
    public static PdfCore.PdfSaveResult SaveAsVisualPdfResult(this OneNoteSection section, string path, OneNoteVisualPdfOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try { return section.ToVisualPdfDocumentResult(options, cancellationToken).SaveResult(path, cancellationToken); }
        catch (OperationCanceledException) { throw; }
        catch (Exception exception) { return PdfCore.PdfSaveResult.FromFailure(path, exception); }
    }

    /// <summary>Converts synchronously, then asynchronously writes a section as visual PDF.</summary>
    public static Task<PdfCore.PdfSaveResult> SaveAsVisualPdfAsync(this OneNoteSection section, string path, OneNoteVisualPdfOptions? options = null, CancellationToken cancellationToken = default) =>
        section.ToVisualPdfDocumentResult(options, cancellationToken).SaveAsync(path, cancellationToken);

    /// <summary>Converts synchronously, then asynchronously writes a section with structured failure evidence.</summary>
    public static async Task<PdfCore.PdfSaveResult> SaveAsVisualPdfResultAsync(this OneNoteSection section, string path, OneNoteVisualPdfOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try { return await section.ToVisualPdfDocumentResult(options, cancellationToken).SaveResultAsync(path, cancellationToken).ConfigureAwait(false); }
        catch (OperationCanceledException) { throw; }
        catch (Exception exception) { return PdfCore.PdfSaveResult.FromFailure(path, exception); }
    }

    /// <summary>Writes a section as visual PDF and throws on failure.</summary>
    public static PdfCore.PdfSaveResult SaveAsVisualPdf(this OneNoteSection section, Stream stream, OneNoteVisualPdfOptions? options = null, CancellationToken cancellationToken = default) =>
        section.ToVisualPdfDocumentResult(options, cancellationToken).Save(stream, cancellationToken);

    /// <summary>Writes a section as visual PDF and captures failure evidence; cancellation still throws.</summary>
    public static PdfCore.PdfSaveResult SaveAsVisualPdfResult(this OneNoteSection section, Stream stream, OneNoteVisualPdfOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try { return section.ToVisualPdfDocumentResult(options, cancellationToken).SaveResult(stream, cancellationToken); }
        catch (OperationCanceledException) { throw; }
        catch (Exception exception) { return PdfCore.PdfSaveResult.FromFailure(null, exception); }
    }

    /// <summary>Converts synchronously, then asynchronously writes a section as visual PDF.</summary>
    public static Task<PdfCore.PdfSaveResult> SaveAsVisualPdfAsync(this OneNoteSection section, Stream stream, OneNoteVisualPdfOptions? options = null, CancellationToken cancellationToken = default) =>
        section.ToVisualPdfDocumentResult(options, cancellationToken).SaveAsync(stream, cancellationToken);

    /// <summary>Converts synchronously, then asynchronously writes a section with structured failure evidence.</summary>
    public static async Task<PdfCore.PdfSaveResult> SaveAsVisualPdfResultAsync(this OneNoteSection section, Stream stream, OneNoteVisualPdfOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try { return await section.ToVisualPdfDocumentResult(options, cancellationToken).SaveResultAsync(stream, cancellationToken).ConfigureAwait(false); }
        catch (OperationCanceledException) { throw; }
        catch (Exception exception) { return PdfCore.PdfSaveResult.FromFailure(null, exception); }
    }

    private static PdfCore.PdfDocumentConversionResult RenderVisualPdf(
        OneNoteSection section,
        OneNoteVisualPdfOptions? options,
        CancellationToken cancellationToken) {
        if (section == null) throw new ArgumentNullException(nameof(section));
        cancellationToken.ThrowIfCancellationRequested();
        IReadOnlyList<OneNotePageReference> pages = OneNotePageTraversal.Flatten(section);
        cancellationToken.ThrowIfCancellationRequested();
        return OneNoteVisualPdfRenderer.Render(section.Name, pages, options, cancellationToken);
    }

    /// <summary>Converts a notebook to a visual PDF document with rendering diagnostics.</summary>
    public static PdfCore.PdfDocumentConversionResult ToVisualPdfDocumentResult(this OneNoteNotebook notebook, OneNoteVisualPdfOptions? options = null, CancellationToken cancellationToken = default) =>
        RenderVisualPdf(notebook, options, cancellationToken);

    /// <summary>Converts a notebook to a visual PDF document.</summary>
    public static PdfCore.PdfDocument ToVisualPdfDocument(this OneNoteNotebook notebook, OneNoteVisualPdfOptions? options = null, CancellationToken cancellationToken = default) =>
        notebook.ToVisualPdfDocumentResult(options, cancellationToken).Value;

    /// <summary>Converts a notebook to serialized visual PDF bytes.</summary>
    public static byte[] ToVisualPdfBytes(this OneNoteNotebook notebook, OneNoteVisualPdfOptions? options = null, CancellationToken cancellationToken = default) =>
        notebook.ToVisualPdfDocumentResult(options, cancellationToken).ToBytes(cancellationToken);

    /// <summary>Writes a notebook as visual PDF and throws on failure.</summary>
    public static PdfCore.PdfSaveResult SaveAsVisualPdf(this OneNoteNotebook notebook, string path, OneNoteVisualPdfOptions? options = null, CancellationToken cancellationToken = default) =>
        notebook.ToVisualPdfDocumentResult(options, cancellationToken).Save(path, cancellationToken);

    /// <summary>Writes a notebook as visual PDF and captures failure evidence; cancellation still throws.</summary>
    public static PdfCore.PdfSaveResult SaveAsVisualPdfResult(this OneNoteNotebook notebook, string path, OneNoteVisualPdfOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try { return notebook.ToVisualPdfDocumentResult(options, cancellationToken).SaveResult(path, cancellationToken); }
        catch (OperationCanceledException) { throw; }
        catch (Exception exception) { return PdfCore.PdfSaveResult.FromFailure(path, exception); }
    }

    /// <summary>Converts synchronously, then asynchronously writes a notebook as visual PDF.</summary>
    public static Task<PdfCore.PdfSaveResult> SaveAsVisualPdfAsync(this OneNoteNotebook notebook, string path, OneNoteVisualPdfOptions? options = null, CancellationToken cancellationToken = default) =>
        notebook.ToVisualPdfDocumentResult(options, cancellationToken).SaveAsync(path, cancellationToken);

    /// <summary>Converts synchronously, then asynchronously writes a notebook with structured failure evidence.</summary>
    public static async Task<PdfCore.PdfSaveResult> SaveAsVisualPdfResultAsync(this OneNoteNotebook notebook, string path, OneNoteVisualPdfOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try { return await notebook.ToVisualPdfDocumentResult(options, cancellationToken).SaveResultAsync(path, cancellationToken).ConfigureAwait(false); }
        catch (OperationCanceledException) { throw; }
        catch (Exception exception) { return PdfCore.PdfSaveResult.FromFailure(path, exception); }
    }

    /// <summary>Writes a notebook as visual PDF and throws on failure.</summary>
    public static PdfCore.PdfSaveResult SaveAsVisualPdf(this OneNoteNotebook notebook, Stream stream, OneNoteVisualPdfOptions? options = null, CancellationToken cancellationToken = default) =>
        notebook.ToVisualPdfDocumentResult(options, cancellationToken).Save(stream, cancellationToken);

    /// <summary>Writes a notebook as visual PDF and captures failure evidence; cancellation still throws.</summary>
    public static PdfCore.PdfSaveResult SaveAsVisualPdfResult(this OneNoteNotebook notebook, Stream stream, OneNoteVisualPdfOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try { return notebook.ToVisualPdfDocumentResult(options, cancellationToken).SaveResult(stream, cancellationToken); }
        catch (OperationCanceledException) { throw; }
        catch (Exception exception) { return PdfCore.PdfSaveResult.FromFailure(null, exception); }
    }

    /// <summary>Converts synchronously, then asynchronously writes a notebook as visual PDF.</summary>
    public static Task<PdfCore.PdfSaveResult> SaveAsVisualPdfAsync(this OneNoteNotebook notebook, Stream stream, OneNoteVisualPdfOptions? options = null, CancellationToken cancellationToken = default) =>
        notebook.ToVisualPdfDocumentResult(options, cancellationToken).SaveAsync(stream, cancellationToken);

    /// <summary>Converts synchronously, then asynchronously writes a notebook with structured failure evidence.</summary>
    public static async Task<PdfCore.PdfSaveResult> SaveAsVisualPdfResultAsync(this OneNoteNotebook notebook, Stream stream, OneNoteVisualPdfOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try { return await notebook.ToVisualPdfDocumentResult(options, cancellationToken).SaveResultAsync(stream, cancellationToken).ConfigureAwait(false); }
        catch (OperationCanceledException) { throw; }
        catch (Exception exception) { return PdfCore.PdfSaveResult.FromFailure(null, exception); }
    }

    private static PdfCore.PdfDocumentConversionResult RenderVisualPdf(
        OneNoteNotebook notebook,
        OneNoteVisualPdfOptions? options,
        CancellationToken cancellationToken) {
        if (notebook == null) throw new ArgumentNullException(nameof(notebook));
        cancellationToken.ThrowIfCancellationRequested();
        IReadOnlyList<OneNotePageReference> pages = OneNotePageTraversal.Flatten(notebook);
        cancellationToken.ThrowIfCancellationRequested();
        return OneNoteVisualPdfRenderer.Render(notebook.Name, pages, options, cancellationToken);
    }

}
