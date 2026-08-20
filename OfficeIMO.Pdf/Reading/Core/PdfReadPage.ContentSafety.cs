using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    internal bool TryGetContentSafetyBackground(PdfTextSpan span, out OfficeColor color, out string evidence) {
        (double width, double height) = GetVisualPageSize();
        Matrix2D transform = GetVisualPageTransform();
        (double left, double top, double right, double bottom) = GetTextVisualBounds(span, height);
        double x = (left + right) / 2D;
        double y = (top + bottom) / 2D;
        double bestPaintOrder = double.NegativeInfinity;
        OfficeColor? best = null;
        bool unresolvedPaint = false;

        foreach (PdfPageVisualPrimitive primitive in GetVisualPrimitives(width, height, transform)) {
            if (primitive.PaintOrder >= span.PaintOrder || !primitive.HasFillPaint ||
                x < primitive.X || x > primitive.X + primitive.Width || y < primitive.Y || y > primitive.Y + primitive.Height) continue;
            if (primitive.PaintOrder < bestPaintOrder) continue;
            bestPaintOrder = primitive.PaintOrder;
            if (primitive.FillColor.HasValue && (primitive.FillOpacity ?? 1D) >= 0.99D) {
                best = primitive.FillColor.Value;
                unresolvedPaint = false;
            } else {
                best = null;
                unresolvedPaint = true;
            }
        }

        foreach (PdfImagePlacement image in GetVisualImagePlacements(height, transform)) {
            double imageTop = height - image.Y - image.Height;
            if (image.PaintOrder >= span.PaintOrder || x < image.X || x > image.X + image.Width || y < imageTop || y > imageTop + image.Height) continue;
            if (image.PaintOrder >= bestPaintOrder) {
                bestPaintOrder = image.PaintOrder;
                best = null;
                unresolvedPaint = true;
            }
        }

        if (unresolvedPaint || HasUnboundedUnsupportedPaint()) {
            color = default;
            evidence = string.Empty;
            return false;
        }
        if (best.HasValue) {
            color = best.Value;
            evidence = "the latest resolved opaque paint beneath the span";
            return true;
        }
        color = OfficeColor.White;
        evidence = "the default unpainted white PDF page background";
        return true;
    }
}
