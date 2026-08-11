using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    private static bool IsInvisibleImagePlacement(
        PdfImagePlacement placement,
        double pageHeight,
        double drawingWidth,
        double drawingHeight) {
        if ((placement.ImageOpacity ?? 1D) <= 0D) return true;
        if (!IsFinite(placement.A) || !IsFinite(placement.B) ||
            !IsFinite(placement.C) || !IsFinite(placement.D) ||
            !IsFinite(placement.E) || !IsFinite(placement.F) ||
            !IsFinite(placement.X) || !IsFinite(placement.Y) ||
            !IsFinite(placement.Width) || !IsFinite(placement.Height) ||
            placement.Width <= 0D || placement.Height <= 0D) return false;
        var geometryBudget = new VisualGeometryBudget();
        var imageTransform = new OfficeTransform(
            placement.A,
            -placement.B,
            placement.C,
            -placement.D,
            placement.E,
            pageHeight - placement.F);
        VisualPath? placementPath = VisualPath.Rectangle(
            0D,
            0D,
            1D,
            1D,
            imageTransform,
            geometryBudget);
        VisualPath? drawingPath = VisualPath.FromClip(
            PdfPageClipPath.Rectangle(0D, 0D, drawingWidth, drawingHeight),
            geometryBudget);
        if (placementPath == null || drawingPath == null || geometryBudget.Exceeded) {
            return false;
        }

        var visiblePaths = new List<VisualPath> { placementPath, drawingPath };
        if (placement.ClipPath.HasValue) {
            VisualPath? clipPath = VisualPath.FromClip(placement.ClipPath.Value, geometryBudget);
            if (clipPath == null || geometryBudget.Exceeded) return false;
            visiblePaths.Add(clipPath);
        }
        return !VisualPath.HasPositiveAreaIntersection(visiblePaths, geometryBudget) && !geometryBudget.Exceeded;
    }
}
