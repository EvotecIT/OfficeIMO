using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public static partial class OfficeDrawingRasterRenderer {
    // Paint belongs to the shape canvas, including any padding. Pull pixel samples
    // back through the shape transform instead of refitting paint to a contour box.
    private static void FillShapeGradientContours(OfficeRasterCanvas canvas, OfficeShape shape,
        double x, double y, double scale, IReadOnlyList<IReadOnlyList<OfficePoint>> contours,
        OfficeLinearGradient? linear, OfficeRadialGradient? radial, OfficeFillRule fillRule) {
        OfficeTransform transform = shape.Transform ?? OfficeTransform.Identity;
        double matrixScale = System.Math.Max(System.Math.Max(System.Math.Abs(transform.M11), System.Math.Abs(transform.M12)),
            System.Math.Max(System.Math.Abs(transform.M21), System.Math.Abs(transform.M22)));
        if (matrixScale == 0) return;
        double a = transform.M11 / matrixScale, b = transform.M12 / matrixScale;
        double c = transform.M21 / matrixScale, d = transform.M22 / matrixScale, determinant = a * d - b * c;
        if (determinant == 0) return;
        canvas.FillContourPaint(contours, fillRule, (px, py) => {
            double destinationX = px / scale - x - transform.OffsetX, destinationY = py / scale - y - transform.OffsetY;
            double localX = ((d * destinationX - c * destinationY) / determinant) / matrixScale / shape.Width;
            double localY = ((a * destinationY - b * destinationX) / determinant) / matrixScale / shape.Height;
            return OfficeRasterCanvas.SampleGradientFill(linear, radial, localX, localY);
        });
    }
}
