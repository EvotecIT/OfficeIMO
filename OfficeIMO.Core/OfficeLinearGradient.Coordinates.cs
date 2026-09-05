using System;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeLinearGradient {
    /// <summary>
    /// Re-expresses the complete gradient color field in another coordinate system.
    /// Transforming endpoints alone is insufficient under non-uniform scale or shear:
    /// equal-color lines must transform with the geometry too.
    /// </summary>
    internal OfficeLinearGradient TransformCoordinates(OfficeTransform transform) {
        if (transform == OfficeTransform.Identity) return this;
        double matrixScale = Math.Max(Math.Max(Math.Abs(transform.M11), Math.Abs(transform.M12)),
            Math.Max(Math.Abs(transform.M21), Math.Abs(transform.M22)));
        if (matrixScale == 0D) {
            throw new ArgumentException("Gradient coordinate transform must be invertible.", nameof(transform));
        }
        double a = transform.M11 / matrixScale;
        double b = transform.M12 / matrixScale;
        double c = transform.M21 / matrixScale;
        double d = transform.M22 / matrixScale;
        double determinant = a * d - b * c;
        if (determinant == 0D) {
            throw new ArgumentException("Gradient coordinate transform must be invertible.", nameof(transform));
        }

        double dx = EndX - StartX;
        double dy = EndY - StartY;
        double lengthSquared = dx * dx + dy * dy;
        // The gradient ratio is a scalar field. Its normal transforms by the
        // inverse transpose, unlike a point or a geometric direction vector.
        double sourceNormalX = dx / lengthSquared;
        double sourceNormalY = dy / lengthSquared;
        double normalX = ((d * sourceNormalX - b * sourceNormalY) / determinant) / matrixScale;
        double normalY = ((a * sourceNormalY - c * sourceNormalX) / determinant) / matrixScale;
        double normalScale = Math.Max(Math.Abs(normalX), Math.Abs(normalY));
        if (normalScale == 0D || double.IsNaN(normalScale) || double.IsInfinity(normalScale)) {
            throw new ArgumentException("Gradient coordinate transform exceeds the representable color field.", nameof(transform));
        }
        normalX /= normalScale;
        normalY /= normalScale;
        double normalSquared = normalX * normalX + normalY * normalY;
        OfficePoint start = transform.TransformPoint(new OfficePoint(StartX, StartY));
        return CreateImported(start.X, start.Y,
            start.X + (normalX / normalSquared) / normalScale,
            start.Y + (normalY / normalSquared) / normalScale,
            Stops);
    }
}
