using System.Collections.Generic;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.OneNote.Pdf;

/// <summary>Visual-preserving PDF entry points backed by the native OneNote Drawing canvas.</summary>
public static partial class OneNoteVisualPdfExtensions {
    private static PdfCore.PdfDocumentConversionResult RenderVisualPdf(
        OneNoteSection section,
        OneNoteVisualPdfExportOptions? options,
        CancellationToken cancellationToken) {
        if (section == null) throw new ArgumentNullException(nameof(section));
        cancellationToken.ThrowIfCancellationRequested();
        IReadOnlyList<OneNotePageReference> pages = OneNotePageTraversal.Flatten(section);
        cancellationToken.ThrowIfCancellationRequested();
        return OneNoteVisualPdfRenderer.Render(section.Name, pages, options, cancellationToken);
    }

    private static PdfCore.PdfDocumentConversionResult RenderVisualPdf(
        OneNoteNotebook notebook,
        OneNoteVisualPdfExportOptions? options,
        CancellationToken cancellationToken) {
        if (notebook == null) throw new ArgumentNullException(nameof(notebook));
        cancellationToken.ThrowIfCancellationRequested();
        IReadOnlyList<OneNotePageReference> pages = OneNotePageTraversal.Flatten(notebook);
        cancellationToken.ThrowIfCancellationRequested();
        return OneNoteVisualPdfRenderer.Render(notebook.Name, pages, options, cancellationToken);
    }

}
