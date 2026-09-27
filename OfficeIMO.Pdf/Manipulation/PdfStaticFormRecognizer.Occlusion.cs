using System.Threading;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfStaticFormRecognizer {
    // PDF fills implicitly close their subpath. Only one exact rectangle proves
    // a rectangular opaque cover; compound paths can have holes under either fill rule.
    private static bool HasExactRectangularFill(PdfPageVisualPrimitive primitive) {
        if (primitive.Kind == PdfPageVisualPrimitiveKind.Rectangle) return true;
        if (primitive.Kind != PdfPageVisualPrimitiveKind.Path) return false;
        return PdfRectanglePathGeometry.IsRectangle(primitive.PathCommands, allowImplicitClose: true);
    }

    // Full containment alone misses an erased side or a broken writing line.
    // Unknown fill geometry is also conservative evidence: its bounds cannot
    // establish that every required outline segment survives the later paint.
    private static bool HasLaterOutlinePaint(IReadOnlyList<PaintArea> filledAreas,
        IReadOnlyList<PdfPageVisualPrimitive> primitives, PdfLogicalPage page,
        PdfPageVisualPrimitive outline, VisualRect candidate, CancellationToken cancellationToken) {
        double padding = Math.Max(0.5D, outline.StrokeWidth * Math.Sqrt(2D)) / 2D;
        VisualRect outer = outline.Kind == PdfPageVisualPrimitiveKind.Line
            ? new VisualRect(Math.Min(outline.X1, outline.X2) - padding,
                Math.Min(outline.Y1, outline.Y2) - padding,
                Math.Max(outline.X1, outline.X2) + padding,
                Math.Max(outline.Y1, outline.Y2) + padding)
            : new VisualRect(candidate.Left - padding, candidate.Top - padding,
                candidate.Right + padding, candidate.Bottom + padding);
        var inner = new VisualRect(candidate.Left + padding, candidate.Top + padding,
            candidate.Right - padding, candidate.Bottom - padding);
        foreach (PaintArea area in filledAreas) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!IsLater(area.PaintOrder, area.ContentOrderKey, outline.PaintOrder, outline.ContentOrderKey)) continue;
            if (IntersectsOutline(area.Bounds)) return true;
        }
        foreach (PdfPageVisualPrimitive paint in primitives) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!paint.HasStrokePaint || paint.StrokeOpacity == 0D ||
                !IsLater(paint.PaintOrder, paint.ContentOrderKey, outline.PaintOrder, outline.ContentOrderKey)) continue;
            if (!TryGetStrokeBounds(paint, out VisualRect bounds, out VisualRect? hollow)) continue;
            if (OutlineOverlap(bounds) - (hollow is VisualRect hole ? OutlineOverlap(hole) : 0D) > 0.000001D) return true;
        }
        foreach (PdfLogicalImage image in page.Images) {
            foreach (PdfImagePlacement placement in image.Placements) {
                cancellationToken.ThrowIfCancellationRequested();
                if (placement.IsHiddenOptionalContent || placement.Opacity <= 0D ||
                    placement.Width <= 0D || placement.Height <= 0D ||
                    !IsLater(placement.PaintOrder, placement.ContentOrderKey, outline.PaintOrder, outline.ContentOrderKey)) continue;
                PdfVisualBounds mapped = page.TransformBoundsToVisual(placement.X, placement.Y,
                    placement.X + placement.Width, placement.Y + placement.Height);
                var bounds = new VisualRect(mapped.Left, mapped.Top, mapped.Right, mapped.Bottom);
                if (placement.Clip is { IsRectangle: true, IsExact: true, ContainsTextClipping: false } clip) {
                    PdfVisualBounds clipped = page.TransformBoundsToVisual(clip.X,
                        page.Height - clip.Y - clip.Height, clip.X + clip.Width, page.Height - clip.Y);
                    bounds = new VisualRect(Math.Max(bounds.Left, clipped.Left), Math.Max(bounds.Top, clipped.Top),
                        Math.Min(bounds.Right, clipped.Right), Math.Min(bounds.Bottom, clipped.Bottom));
                }
                if (placement.Clip is { IsRectangle: false, IsExact: true, ContainsTextClipping: false } path &&
                    PdfPageClipPath.TryCreatePath(path.Commands, path.FillRule, out PdfPageClipPath exactClip)) {
                    PdfPageRectangle user = page.MapVisualRectangleToUserSpace(outer.Left, outer.Top, outer.Right, outer.Bottom);
                    PdfPageClipPath outlineClip = PdfPageClipPath.Rectangle(user.Left, page.Height - user.Top, user.Width, user.Height);
                    if (exactClip.CanProveNoPositiveAreaIntersection(outlineClip)) continue;
                }
                if (IntersectsOutline(bounds)) return true;
            }
        }
        return false;

        bool IntersectsOutline(VisualRect bounds) => OutlineOverlap(bounds) > 0.000001D;

        double OutlineOverlap(VisualRect bounds) {
            double overlap = OverlapArea(bounds, outer);
            if (outline.Kind != PdfPageVisualPrimitiveKind.Line) overlap -= OverlapArea(bounds, inner);
            return overlap;
        }
    }

    private static bool TryGetStrokeBounds(PdfPageVisualPrimitive paint, out VisualRect bounds, out VisualRect? hollow) {
        bounds = default;
        hollow = null;
        if (!paint.HasStrokePaint || paint.StrokeOpacity == 0D) return false;
        double strokePadding = Math.Max(0.5D, paint.StrokeWidth * Math.Sqrt(2D)) / 2D;
        bounds = paint.Kind == PdfPageVisualPrimitiveKind.Line
            ? new VisualRect(Math.Min(paint.X1, paint.X2) - strokePadding,
                Math.Min(paint.Y1, paint.Y2) - strokePadding,
                Math.Max(paint.X1, paint.X2) + strokePadding,
                Math.Max(paint.Y1, paint.Y2) + strokePadding)
            : new VisualRect(paint.X - strokePadding, paint.Y - strokePadding,
                paint.X + paint.Width + strokePadding, paint.Y + paint.Height + strokePadding);
        hollow = paint.Kind == PdfPageVisualPrimitiveKind.Rectangle
            ? new VisualRect(paint.X + strokePadding, paint.Y + strokePadding,
                paint.X + paint.Width - strokePadding, paint.Y + paint.Height - strokePadding) : null;
        if (paint.ClipPath is { IsRectangle: true, IsExact: true, ContainsTextClipping: false } clip) {
            bounds = new VisualRect(Math.Max(bounds.Left, clip.X), Math.Max(bounds.Top, clip.Y),
                Math.Min(bounds.Right, clip.X + clip.Width), Math.Min(bounds.Bottom, clip.Y + clip.Height));
            if (hollow is VisualRect empty) hollow = new VisualRect(Math.Max(empty.Left, clip.X), Math.Max(empty.Top, clip.Y),
                Math.Min(empty.Right, clip.X + clip.Width), Math.Min(empty.Bottom, clip.Y + clip.Height));
        }
        return bounds.Area > 0D;
    }
}
