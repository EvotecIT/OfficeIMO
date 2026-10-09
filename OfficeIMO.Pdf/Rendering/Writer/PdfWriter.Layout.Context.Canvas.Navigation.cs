using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private void RenderCanvasOutline(PdfCanvasOutlineItem item) {
            EnsurePage();
            OfficePoint point = _canvasEffectToPage.TransformPoint(new OfficePoint(item.X, currentOpts.PageHeight - item.Y));
            currentPage!.Bookmarks.Add(new PageBookmark {
                Level = item.Level,
                Title = item.Title,
                Y = point.Y,
                OutlineState = item.State,
                DocumentOrder = item.DocumentOrder, X = point.X, Uri = item.Uri
            });
            pageDirty = true;
        }

        private void RenderCanvasNamedDestination(PdfCanvasNamedDestinationItem item) {
            EnsurePage();
            OfficePoint point = _canvasEffectToPage.TransformPoint(
                new OfficePoint(item.X, currentOpts.PageHeight - item.Y));
            currentPage!.NamedDestinations.Add(new PageNamedDestination { Name = item.Name, X = point.X, Y = point.Y });
        }

    }
}
