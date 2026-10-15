using System.Threading;

namespace OfficeIMO.Pdf;

/// <summary>Qualifies optional source projections before consumers infer content absence or loss.</summary>
internal static class PdfLogicalContentProjection {
    internal static void RequireComplete(PdfDocumentReadResult document, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!document.IncludesImages || !document.IncludesFormWidgets) {
            throw IncompleteProjection();
        }
        RequireComplete(document.Pages, cancellationToken);
    }

    internal static void RequireComplete(IReadOnlyList<PdfLogicalPage> pages, CancellationToken cancellationToken = default) {
        for (int i = 0; i < pages.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfLogicalPage page = pages[i];
            if (!page.IncludesImages || !page.IncludesFormWidgets || !page.HasVisualPrimitiveCounts) {
                throw IncompleteProjection();
            }
        }
    }

    private static InvalidOperationException IncompleteProjection() => new InvalidOperationException(
        "Content omission analysis and conversion require inspected image, form-widget, and visual primitive projections. " +
        "Read the PDF with IncludeImages and IncludeFormWidgets enabled; an unrequested projection cannot establish content absence.");
}
