using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    internal static bool IsInvisibleImagePlacement(
        PdfImagePlacement placement,
        double pageHeight,
        double drawingWidth,
        double drawingHeight,
        VisualGeometryBudget? geometryBudget = null) {
        if ((placement.ImageOpacity ?? 1D) <= 0D) return true;
        if (!IsFinite(placement.A) || !IsFinite(placement.B) ||
            !IsFinite(placement.C) || !IsFinite(placement.D) ||
            !IsFinite(placement.E) || !IsFinite(placement.F) ||
            !IsFinite(placement.X) || !IsFinite(placement.Y) ||
            !IsFinite(placement.Width) || !IsFinite(placement.Height) ||
            placement.Width <= 0D || placement.Height <= 0D) return false;
        geometryBudget ??= new VisualGeometryBudget();
        var imageTransform = new OfficeTransform(
            placement.A,
            -placement.B,
            placement.C,
            -placement.D,
            placement.E,
            pageHeight - placement.F);
        OfficePoint[] imageCorners = {
            imageTransform.TransformPoint(new OfficePoint(0D, 0D)),
            imageTransform.TransformPoint(new OfficePoint(1D, 0D)),
            imageTransform.TransformPoint(new OfficePoint(1D, 1D)),
            imageTransform.TransformPoint(new OfficePoint(0D, 1D))
        };
        OfficePathCommand[] imageCommands = {
            OfficePathCommand.MoveTo(imageCorners[0]),
            OfficePathCommand.LineTo(imageCorners[1]),
            OfficePathCommand.LineTo(imageCorners[2]),
            OfficePathCommand.LineTo(imageCorners[3]),
            OfficePathCommand.Close()
        };
        if (PdfPageClipPath.TryCreatePath(imageCommands, OfficeFillRule.NonZero, out PdfPageClipPath imageClip)) {
            PdfPageClipPath pageCandidate = PdfPageClipPath.ResolveActiveClip(
                imageClip,
                PdfPageClipPath.Rectangle(0D, 0D, drawingWidth, drawingHeight));
            if (pageCandidate.IsExact && (pageCandidate.Width <= 0D || pageCandidate.Height <= 0D)) return true;
            PdfPageClipPath exactCandidate = pageCandidate;
            if (placement.ClipPath.HasValue) {
                if (imageClip.CanProveNoPositiveAreaIntersection(placement.ClipPath.Value, geometryBudget)) return true;
                exactCandidate = PdfPageClipPath.ResolveActiveClip(exactCandidate, placement.ClipPath.Value);
                if (pageCandidate.CanProveNoPositiveAreaIntersection(placement.ClipPath.Value, geometryBudget)) return true;
            }
            if ((!placement.ClipPath.HasValue || placement.ClipPath.Value.CanProveExactIntersection) &&
                exactCandidate.IsExact) return exactCandidate.Width <= 0D || exactCandidate.Height <= 0D;
        }
        // A sampled miss cannot prove that an inexact, curved, or concave clip has no
        // positive-area overlap. Retain the image so the later strict gate can either
        // project it exactly or fail closed to native Type 3 substitution.
        return false;
    }
}
