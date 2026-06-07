namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static void ValidateOpenAction(PdfOpenActionOptions openAction, int pageCount) {
        if (openAction.PageNumber > pageCount) {
            throw new InvalidOperationException("PDF open-action page number cannot exceed the generated page count.");
        }
    }

    private static (double Left, double Bottom, double Right, double Top) ResolveOpenActionDestinationCoordinates(PdfOpenActionOptions openAction, LayoutResult.Page page) {
        double left = openAction.DestinationLeft ?? 0d;
        double bottom = openAction.DestinationBottom ?? 0d;
        double right = openAction.DestinationRight ?? page.Options.PageWidth;
        double top = openAction.DestinationTop ?? page.Options.PageHeight;

        if (right <= left) {
            throw new InvalidOperationException("PDF open-action destination rectangle right coordinate must be greater than left coordinate.");
        }

        if (top <= bottom) {
            throw new InvalidOperationException("PDF open-action destination rectangle top coordinate must be greater than bottom coordinate.");
        }

        return (left, bottom, right, top);
    }
}
