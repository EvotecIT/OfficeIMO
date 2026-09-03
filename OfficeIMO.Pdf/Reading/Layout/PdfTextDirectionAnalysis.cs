using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static class PdfTextDirectionAnalysis {
    internal static PdfReadingDirection Resolve(
        PdfReadingDirection requested,
        IEnumerable<string> textInSourceOrder) {
        if (requested != PdfReadingDirection.Auto) return requested;
        foreach (string text in textInSourceOrder) {
            OfficeTextDirection direction = OfficeTextElements.ResolveBaseDirection(text);
            if (direction == OfficeTextDirection.RightToLeft) return PdfReadingDirection.RightToLeft;
            if (direction == OfficeTextDirection.LeftToRight) return PdfReadingDirection.LeftToRight;
        }
        return PdfReadingDirection.LeftToRight;
    }
}
