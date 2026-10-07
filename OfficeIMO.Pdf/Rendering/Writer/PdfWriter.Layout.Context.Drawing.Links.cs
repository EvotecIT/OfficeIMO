using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private void DrawDrawingLinkAt(OfficeDrawingLink link, double originX, double originTopY) {
            if (_suppressDrawingLinks) return;
            var annotation = new LinkAnnotation {
                X1 = originX + link.X,
                X2 = originX + link.X + link.Width,
                Y1 = originTopY - link.Y - link.Height,
                Y2 = originTopY - link.Y,
                Contents = link.AlternativeText
            };
            if (link.Uri[0] == '#') annotation.DestinationName = Uri.UnescapeDataString(link.Uri).Substring(1);
            else annotation.Uri = link.Uri;
            currentPage!.Annotations.Add(annotation);
            pageDirty = true;
        }
    }
}
