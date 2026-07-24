using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// Dependency-free 2D affine transform in a local top-left coordinate space.
/// </summary>
public readonly struct OfficeTransform : IEquatable<OfficeTransform> {
    /// <summary>Identity transform.</summary>
    public static OfficeTransform Identity => new OfficeTransform(1, 0, 0, 1, 0, 0);

    /// <summary>Horizontal scale/rotation component.</summary>
    public double M11 { get; }

    /// <summary>Vertical shear/rotation component.</summary>
    public double M12 { get; }

    /// <summary>Horizontal shear/rotation component.</summary>
    public double M21 { get; }

    /// <summary>Vertical scale/rotation component.</summary>
    public double M22 { get; }

    /// <summary>Horizontal translation component.</summary>
    public double OffsetX { get; }

    /// <summary>Vertical translation component.</summary>
    public double OffsetY { get; }

    /// <summary>Creates an affine transform.</summary>
    public OfficeTransform(double m11, double m12, double m21, double m22, double offsetX, double offsetY) {
        ValidateFinite(m11, nameof(m11));
        ValidateFinite(m12, nameof(m12));
        ValidateFinite(m21, nameof(m21));
        ValidateFinite(m22, nameof(m22));
        ValidateFinite(offsetX, nameof(offsetX));
        ValidateFinite(offsetY, nameof(offsetY));

        M11 = m11;
        M12 = m12;
        M21 = m21;
        M22 = m22;
        OffsetX = offsetX;
        OffsetY = offsetY;
    }

    /// <summary>Creates a translation transform.</summary>
    public static OfficeTransform Translate(double offsetX, double offsetY) => new OfficeTransform(1, 0, 0, 1, offsetX, offsetY);

    /// <summary>Creates a scale transform around the local origin.</summary>
    public static OfficeTransform Scale(double scaleX, double scaleY) => new OfficeTransform(scaleX, 0, 0, scaleY, 0, 0);

    /// <summary>Creates a clockwise rotation transform around the local origin.</summary>
    public static OfficeTransform RotateDegrees(double degrees) {
        ValidateFinite(degrees, nameof(degrees));

        double normalizedDegrees = Math.IEEERemainder(degrees, 360D);
        double radians = normalizedDegrees * Math.PI / 180D;
        double cos = NormalizeZero(Math.Cos(radians));
        double sin = NormalizeZero(Math.Sin(radians));
        return new OfficeTransform(cos, sin, -sin, cos, 0, 0);
    }

    /// <summary>Creates a clockwise rotation transform around a local point.</summary>
    public static OfficeTransform RotateDegrees(double degrees, double centerX, double centerY) {
        ValidateFinite(centerX, nameof(centerX));
        ValidateFinite(centerY, nameof(centerY));

        return Translate(-centerX, -centerY)
            .Then(RotateDegrees(degrees))
            .Then(Translate(centerX, centerY));
    }

    /// <summary>
    /// Returns a transform that applies this transform first, then the supplied transform.
    /// </summary>
    public OfficeTransform Then(OfficeTransform next) {
        return new OfficeTransform(
            next.M11 * M11 + next.M21 * M12,
            next.M12 * M11 + next.M22 * M12,
            next.M11 * M21 + next.M21 * M22,
            next.M12 * M21 + next.M22 * M22,
            next.M11 * OffsetX + next.M21 * OffsetY + next.OffsetX,
            next.M12 * OffsetX + next.M22 * OffsetY + next.OffsetY);
    }

    /// <summary>
    /// Returns the inverse transform.
    /// </summary>
    /// <exception cref="InvalidOperationException">Thrown when the transform cannot be inverted.</exception>
    public OfficeTransform Invert() {
        if (!TryInvert(out OfficeTransform inverse)) {
            throw new InvalidOperationException("Transform cannot be inverted because its determinant is zero.");
        }

        return inverse;
    }

    /// <summary>
    /// Attempts to create the inverse transform.
    /// </summary>
    /// <param name="inverse">Inverse transform when inversion succeeds.</param>
    /// <returns>True when the transform has a non-zero determinant.</returns>
    public bool TryInvert(out OfficeTransform inverse) {
        double determinant = (M11 * M22) - (M21 * M12);
        if (Math.Abs(determinant) < 0.000000000001D) {
            inverse = default;
            return false;
        }

        inverse = new OfficeTransform(
            NormalizeZero(M22 / determinant),
            NormalizeZero(-M12 / determinant),
            NormalizeZero(-M21 / determinant),
            NormalizeZero(M11 / determinant),
            NormalizeZero(((M21 * OffsetY) - (M22 * OffsetX)) / determinant),
            NormalizeZero(((M12 * OffsetX) - (M11 * OffsetY)) / determinant));
        return true;
    }

    /// <summary>Transforms a point in local top-left coordinates.</summary>
    public OfficePoint TransformPoint(OfficePoint point) {
        double x = M11 * point.X + M21 * point.Y + OffsetX;
        double y = M12 * point.X + M22 * point.Y + OffsetY;
        return new OfficePoint(x, y);
    }

    /// <summary>
    /// Calculates axis-aligned bounds for a transformed rectangle.
    /// </summary>
    /// <param name="x">Rectangle left coordinate.</param>
    /// <param name="y">Rectangle top coordinate.</param>
    /// <param name="width">Rectangle width.</param>
    /// <param name="height">Rectangle height.</param>
    /// <returns>Axis-aligned bounds of the transformed rectangle.</returns>
    public (double Left, double Top, double Right, double Bottom) TransformRectangleBounds(double x, double y, double width, double height) {
        ValidateFinite(x, nameof(x));
        ValidateFinite(y, nameof(y));
        ValidateFinite(width, nameof(width));
        ValidateFinite(height, nameof(height));

        OfficePoint topLeft = TransformPoint(new OfficePoint(x, y));
        OfficePoint topRight = TransformPoint(new OfficePoint(x + width, y));
        OfficePoint bottomRight = TransformPoint(new OfficePoint(x + width, y + height));
        OfficePoint bottomLeft = TransformPoint(new OfficePoint(x, y + height));

        return (
            Math.Min(Math.Min(topLeft.X, topRight.X), Math.Min(bottomRight.X, bottomLeft.X)),
            Math.Min(Math.Min(topLeft.Y, topRight.Y), Math.Min(bottomRight.Y, bottomLeft.Y)),
            Math.Max(Math.Max(topLeft.X, topRight.X), Math.Max(bottomRight.X, bottomLeft.X)),
            Math.Max(Math.Max(topLeft.Y, topRight.Y), Math.Max(bottomRight.Y, bottomLeft.Y)));
    }

    /// <inheritdoc />
    public bool Equals(OfficeTransform other) {
        return M11.Equals(other.M11) &&
               M12.Equals(other.M12) &&
               M21.Equals(other.M21) &&
               M22.Equals(other.M22) &&
               OffsetX.Equals(other.OffsetX) &&
               OffsetY.Equals(other.OffsetY);
    }

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is OfficeTransform other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() {
        unchecked {
            int hash = M11.GetHashCode();
            hash = (hash * 397) ^ M12.GetHashCode();
            hash = (hash * 397) ^ M21.GetHashCode();
            hash = (hash * 397) ^ M22.GetHashCode();
            hash = (hash * 397) ^ OffsetX.GetHashCode();
            hash = (hash * 397) ^ OffsetY.GetHashCode();
            return hash;
        }
    }

    /// <summary>Returns true when two transforms are equal.</summary>
    public static bool operator ==(OfficeTransform left, OfficeTransform right) => left.Equals(right);

    /// <summary>Returns true when two transforms are not equal.</summary>
    public static bool operator !=(OfficeTransform left, OfficeTransform right) => !left.Equals(right);

    private static double NormalizeZero(double value) => Math.Abs(value) < 0.000000000001D ? 0D : value;

    private static void ValidateFinite(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value)) {
            throw new ArgumentOutOfRangeException(paramName, "Transform values must be finite numbers.");
        }
    }
}
