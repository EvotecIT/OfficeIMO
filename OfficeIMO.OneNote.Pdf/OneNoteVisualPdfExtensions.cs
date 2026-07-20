using System.Collections.Generic;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.OneNote.Pdf;

/// <summary>Visual-preserving PDF entry points backed by the native OneNote Drawing canvas.</summary>
public static class OneNoteVisualPdfExtensions {
    /// <summary>Converts a section to a first-party PDF result with rendering diagnostics.</summary>
    public static PdfCore.PdfDocumentConversionResult ToVisualPdfDocumentResult(this OneNoteSection section, OneNoteVisualPdfOptions? options = null) {
        if (section == null) throw new ArgumentNullException(nameof(section));
        return OneNoteVisualPdfRenderer.Render(section.Name, OneNotePageTraversal.Flatten(section), options);
    }

    /// <summary>Converts a notebook to a first-party PDF result with rendering diagnostics.</summary>
    public static PdfCore.PdfDocumentConversionResult ToVisualPdfDocumentResult(this OneNoteNotebook notebook, OneNoteVisualPdfOptions? options = null) {
        if (notebook == null) throw new ArgumentNullException(nameof(notebook));
        return OneNoteVisualPdfRenderer.Render(notebook.Name, OneNotePageTraversal.Flatten(notebook), options);
    }

    /// <summary>Converts a section to a first-party visual PDF document.</summary>
    public static PdfCore.PdfDocument ToVisualPdfDocument(this OneNoteSection section, OneNoteVisualPdfOptions? options = null) =>
        section.ToVisualPdfDocumentResult(options).Value;

    /// <summary>Converts a notebook to a first-party visual PDF document.</summary>
    public static PdfCore.PdfDocument ToVisualPdfDocument(this OneNoteNotebook notebook, OneNoteVisualPdfOptions? options = null) =>
        notebook.ToVisualPdfDocumentResult(options).Value;

    /// <summary>Converts a section to visual PDF bytes.</summary>
    public static byte[] ToVisualPdf(this OneNoteSection section, OneNoteVisualPdfOptions? options = null) =>
        section.ToVisualPdfDocument(options).ToBytes();

    /// <summary>Converts a notebook to visual PDF bytes.</summary>
    public static byte[] ToVisualPdf(this OneNoteNotebook notebook, OneNoteVisualPdfOptions? options = null) =>
        notebook.ToVisualPdfDocument(options).ToBytes();

    /// <summary>Saves a section as a visual PDF and returns conversion diagnostics.</summary>
    public static PdfCore.PdfSaveResult SaveAsVisualPdf(this OneNoteSection section, string path, OneNoteVisualPdfOptions? options = null) =>
        section.ToVisualPdfDocumentResult(options).Save(path);

    /// <summary>Saves a notebook as a visual PDF and returns conversion diagnostics.</summary>
    public static PdfCore.PdfSaveResult SaveAsVisualPdf(this OneNoteNotebook notebook, string path, OneNoteVisualPdfOptions? options = null) =>
        notebook.ToVisualPdfDocumentResult(options).Save(path);

    /// <summary>Writes a section as a visual PDF to a caller-owned stream.</summary>
    public static PdfCore.PdfSaveResult SaveAsVisualPdf(this OneNoteSection section, Stream stream, OneNoteVisualPdfOptions? options = null) =>
        section.ToVisualPdfDocumentResult(options).Save(stream);

    /// <summary>Writes a notebook as a visual PDF to a caller-owned stream.</summary>
    public static PdfCore.PdfSaveResult SaveAsVisualPdf(this OneNoteNotebook notebook, Stream stream, OneNoteVisualPdfOptions? options = null) =>
        notebook.ToVisualPdfDocumentResult(options).Save(stream);

    /// <summary>Asynchronously saves a section as a visual PDF.</summary>
    public static Task<PdfCore.PdfSaveResult> SaveAsVisualPdfAsync(this OneNoteSection section, string path, OneNoteVisualPdfOptions? options = null, CancellationToken cancellationToken = default) =>
        RenderVisualPdf(section, options, cancellationToken).SaveAsync(path, cancellationToken);

    /// <summary>Asynchronously saves a notebook as a visual PDF.</summary>
    public static Task<PdfCore.PdfSaveResult> SaveAsVisualPdfAsync(this OneNoteNotebook notebook, string path, OneNoteVisualPdfOptions? options = null, CancellationToken cancellationToken = default) =>
        RenderVisualPdf(notebook, options, cancellationToken).SaveAsync(path, cancellationToken);

    /// <summary>Asynchronously writes a section as a visual PDF to a caller-owned stream.</summary>
    public static Task<PdfCore.PdfSaveResult> SaveAsVisualPdfAsync(this OneNoteSection section, Stream stream, OneNoteVisualPdfOptions? options = null, CancellationToken cancellationToken = default) =>
        RenderVisualPdf(section, options, cancellationToken).SaveAsync(stream, cancellationToken);

    /// <summary>Asynchronously writes a notebook as a visual PDF to a caller-owned stream.</summary>
    public static Task<PdfCore.PdfSaveResult> SaveAsVisualPdfAsync(this OneNoteNotebook notebook, Stream stream, OneNoteVisualPdfOptions? options = null, CancellationToken cancellationToken = default) =>
        RenderVisualPdf(notebook, options, cancellationToken).SaveAsync(stream, cancellationToken);

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
