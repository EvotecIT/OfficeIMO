namespace OfficeIMO.Pdf;

internal static partial class PdfStaticFormRecognizer {
    // The canonical effect resolver may scan the entire timeline. Charge its
    // conservative upper bound before lookup, including pages with no candidates.
    private static void ChargeEffectLookup(int transitionCount, ref long work, int maximumWork,
        System.Threading.CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        work = checked(work + transitionCount + 1L);
        if (work > maximumWork) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.UnderstandingArtifacts, maximumWork, work);
        }
    }

    // Image geometry does not prove its rendered color. Native labels over an image
    // need separate visibility evidence, such as caller-supplied positioned OCR.
    private static bool HasUnprovenImageBackdrop(PdfLogicalPage page,
        PdfImagePlacement placement, VisualRect label, double paintOrder,
        PdfContentOrderKey? contentOrderKey) {
        if (!IsLater(paintOrder, contentOrderKey, placement.PaintOrder, placement.ContentOrderKey) ||
            placement.IsHiddenOptionalContent || placement.Opacity <= 0D ||
            placement.Width <= 0D || placement.Height <= 0D) return false;
        PdfVisualBounds mapped = page.TransformBoundsToVisual(placement.X, placement.Y,
            placement.X + placement.Width, placement.Y + placement.Height);
        var visible = new VisualRect(mapped.Left, mapped.Top, mapped.Right, mapped.Bottom);
        if (placement.Clip is { IsRectangle: true, IsExact: true, ContainsTextClipping: false } clip) {
            PdfVisualBounds clipped = page.TransformBoundsToVisual(clip.X,
                page.Height - clip.Y - clip.Height, clip.X + clip.Width, page.Height - clip.Y);
            visible = new VisualRect(Math.Max(visible.Left, clipped.Left), Math.Max(visible.Top, clipped.Top),
                Math.Min(visible.Right, clipped.Right), Math.Min(visible.Bottom, clipped.Bottom));
        }
        return OverlapArea(visible, label) > 0D;
    }
}
