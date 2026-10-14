using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficeDrawingRasterRenderer {
    // Only trim trailing transparent storage on the existing source texel grid.
    // Drawing dimensions/elements, layer placement, opacity and the public scene stay unchanged.
    private static bool TryResolveFilledPrimitiveEffectPixelBounds(
        OfficeRasterCanvas canvas, OfficeDrawingEffectGroup effect, EffectTextLayerPlan layer,
        OfficeTransform pixelTransform, double fullWidth, double fullHeight,
        CancellationToken cancellationToken, out int widthPixels, out int heightPixels) {
        widthPixels = 0; heightPixels = 0;
        if (effect.BlendMode != OfficeBlendMode.Normal || effect.SoftMask != null ||
            effect.UnfilteredGeometryBounds.HasValue || effect.IsSvgMarkerPaint || effect.HasCompleteLocalPaintBounds ||
            layer.Left != 0D || layer.ParagraphInk.Count != 0 || layer.Drawing.Elements.Count == 0 ||
            canvas.CoordinateScaleX != 1D || canvas.CoordinateScaleY != 1D ||
            layer.ScaleX != layer.ScaleY || !IsFiniteEffectAllocation(layer.ScaleX) || layer.ScaleX <= 0D ||
            pixelTransform.M11 != 1D || pixelTransform.M12 != 0D || pixelTransform.M21 != 0D || pixelTransform.M22 != 1D ||
            !IsExactEffectPixelTranslation(pixelTransform.OffsetX) || !IsExactEffectPixelTranslation(pixelTransform.OffsetY) ||
            !IsFiniteEffectAllocation(fullWidth) || !IsFiniteEffectAllocation(fullHeight)) return false;
        int maximumDimension = OfficeRasterImageEncoder.GetMaximumDimension(OfficeImageExportFormat.Png);
        // Preserve the original encoder-dimension rejection, rather than making an oversized
        // authored surface admissible merely because its paint happens to be small.
        if (fullWidth > maximumDimension || fullHeight > maximumDimension) return false;

        double scale = layer.ScaleX;
        double right = 0D, bottom = 0D;
        var inspection = new SamplingInspectionContext(cancellationToken);
        foreach (OfficeDrawingElement element in layer.Drawing.Elements) {
            if (!inspection.TryConsume() || element is not OfficeDrawingShape positioned) return false;
            OfficeShape shape = positioned.Shape;
            if (shape.Kind != OfficeShapeKind.Rectangle && shape.Kind != OfficeShapeKind.RoundedRectangle && shape.Kind != OfficeShapeKind.Ellipse)
                return false;
            if (shape.Shadow != null || shape.Glow != null || shape.ClipPath != null || HasNonIdentityTransform(shape.Transform) ||
                shape.StrokeWidth != 0D || !IsFiniteEffectOpacity(shape.FillOpacity) || !IsFiniteEffectOpacity(shape.StrokeOpacity) ||
                (!shape.FillColor.HasValue && shape.FillGradient == null && shape.FillRadialGradient == null) ||
                !IsFiniteEffectAllocation(shape.Width) || !IsFiniteEffectAllocation(shape.Height) || shape.Width <= 0D || shape.Height <= 0D ||
                !IsFiniteEffectAllocation(positioned.X) || !IsFiniteEffectAllocation(positioned.Y) ||
                (shape.Kind == OfficeShapeKind.RoundedRectangle && (!IsFiniteEffectAllocation(shape.CornerRadius) || shape.CornerRadius < 0D)))
                return false;

            // Match RenderShapeGeometry's physical arithmetic and exact contour factories.
            // Using (X + Width) * scale instead could cross a pixel ceiling by an ulp.
            double x = positioned.X * scale, y = positioned.Y * scale;
            double width = shape.Width * scale, height = shape.Height * scale;
            IReadOnlyList<OfficePoint> contour = shape.Kind == OfficeShapeKind.Rectangle
                ? CreateRectangleContour(x, y, width, height)
                : OffsetPoints(shape.Kind == OfficeShapeKind.Ellipse
                    ? CreateEllipseContour(width, height, 1D)
                    : CreateRoundedRectangleContour(width, height, shape.CornerRadius * scale, 1D), x, y, 1D);
            foreach (OfficePoint point in contour) {
                if (!inspection.TryConsume() || !IsFiniteEffectAllocation(point.X) || !IsFiniteEffectAllocation(point.Y)) return false;
                right = Math.Max(right, point.X);
                bottom = Math.Max(bottom, point.Y);
            }
            // The union cannot shrink after it already requires the full surface.
            if (right >= fullWidth && bottom >= fullHeight) return false;
        }

        // FillContourPaint writes through ceil(max)-1 and clips negative paint at zero.
        // Keep that original leading boundary; do not translate or add a guessed halo.
        widthPixels = (int)Math.Min(fullWidth, Math.Ceiling(right));
        heightPixels = (int)Math.Min(fullHeight, Math.Ceiling(bottom));
        if (widthPixels <= 0 || heightPixels <= 0 || (widthPixels == fullWidth && heightPixels == fullHeight)) return false;
        return true;
    }

    private static bool IsFiniteEffectAllocation(double value) => !double.IsNaN(value) && !double.IsInfinity(value);

    private static bool IsFiniteEffectOpacity(double? value) => !value.HasValue || IsFiniteEffectAllocation(value.Value);

    private static bool IsExactEffectPixelTranslation(double value) =>
        IsFiniteEffectAllocation(value) && Math.Abs(value) <= int.MaxValue && value == Math.Truncate(value);
}
