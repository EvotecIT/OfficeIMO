using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// Shared text placement helpers for renderers that already translated document-specific alignment values.
/// </summary>
public static class OfficeTextPlacement {
    /// <summary>
    /// Resolves the horizontal text anchor inside a left-based rectangle.
    /// </summary>
    /// <param name="left">Left edge of the available rectangle.</param>
    /// <param name="availableWidth">Available rectangle width.</param>
    /// <param name="alignment">Horizontal text alignment.</param>
    /// <returns>The anchor X coordinate expected by anchored text drawing.</returns>
    public static double ResolveAnchorX(double left, double availableWidth, OfficeTextAlignment alignment) {
        double width = NormalizeNonNegative(availableWidth);
        if (alignment == OfficeTextAlignment.Right) {
            return left + width;
        }

        if (alignment == OfficeTextAlignment.Center) {
            return left + (width / 2D);
        }

        return left;
    }

    /// <summary>
    /// Resolves the horizontal text anchor inside a center-based rectangle.
    /// </summary>
    /// <param name="centerX">Center X of the available rectangle.</param>
    /// <param name="availableWidth">Available rectangle width.</param>
    /// <param name="alignment">Horizontal text alignment.</param>
    /// <returns>The anchor X coordinate expected by anchored text drawing.</returns>
    public static double ResolveAnchorXFromCenter(double centerX, double availableWidth, OfficeTextAlignment alignment) {
        if (!IsFinitePositive(availableWidth)) {
            return centerX;
        }

        if (alignment == OfficeTextAlignment.Left) {
            return centerX - (availableWidth / 2D);
        }

        if (alignment == OfficeTextAlignment.Right) {
            return centerX + (availableWidth / 2D);
        }

        return centerX;
    }

    /// <summary>
    /// Resolves the left coordinate of measured text from an anchor and alignment.
    /// </summary>
    /// <param name="anchorX">Text anchor X coordinate.</param>
    /// <param name="textWidth">Measured text width.</param>
    /// <param name="alignment">Horizontal text alignment.</param>
    /// <returns>The left coordinate where measured text should begin.</returns>
    public static double ResolveLeftFromAnchor(double anchorX, double textWidth, OfficeTextAlignment alignment) {
        double width = NormalizeNonNegative(textWidth);
        if (alignment == OfficeTextAlignment.Right) {
            return anchorX - width;
        }

        if (alignment == OfficeTextAlignment.Center) {
            return anchorX - (width / 2D);
        }

        return anchorX;
    }

    /// <summary>
    /// Resolves the left coordinate of measured text inside a left-based rectangle.
    /// </summary>
    /// <param name="left">Left edge of the available rectangle.</param>
    /// <param name="availableWidth">Available rectangle width.</param>
    /// <param name="textWidth">Measured text width.</param>
    /// <param name="alignment">Horizontal text alignment.</param>
    /// <returns>The left coordinate where measured text should begin.</returns>
    public static double ResolveLineLeft(double left, double availableWidth, double textWidth, OfficeTextAlignment alignment) {
        double width = NormalizeNonNegative(availableWidth);
        double measured = NormalizeNonNegative(textWidth);
        if (alignment == OfficeTextAlignment.Right) {
            return left + Math.Max(0D, width - measured);
        }

        if (alignment == OfficeTextAlignment.Center) {
            return left + Math.Max(0D, (width - measured) / 2D);
        }

        return left;
    }

    /// <summary>
    /// Resolves the top coordinate of measured text inside a top-based rectangle.
    /// </summary>
    /// <param name="top">Top edge of the available rectangle.</param>
    /// <param name="availableHeight">Available rectangle height.</param>
    /// <param name="textHeight">Measured text height.</param>
    /// <param name="alignment">Vertical text alignment.</param>
    /// <returns>The top coordinate where measured text should begin.</returns>
    public static double ResolveTop(double top, double availableHeight, double textHeight, OfficeTextVerticalAlignment alignment) {
        double height = NormalizeNonNegative(availableHeight);
        double measured = NormalizeNonNegative(textHeight);
        if (alignment == OfficeTextVerticalAlignment.Center) {
            return top + Math.Max(0D, (height - measured) / 2D);
        }

        if (alignment == OfficeTextVerticalAlignment.Top) {
            return top;
        }

        return top + Math.Max(0D, height - measured);
    }

    /// <summary>
    /// Resolves the top coordinate of measured text inside a center-based rectangle.
    /// </summary>
    /// <param name="centerY">Center Y of the available rectangle.</param>
    /// <param name="availableHeight">Available rectangle height.</param>
    /// <param name="textHeight">Measured text height.</param>
    /// <param name="alignment">Vertical text alignment.</param>
    /// <returns>The top coordinate where measured text should begin.</returns>
    public static double ResolveTopFromCenter(double centerY, double availableHeight, double textHeight, OfficeTextVerticalAlignment alignment) {
        double measured = NormalizeNonNegative(textHeight);
        if (!IsFinitePositive(availableHeight)) {
            return centerY - (measured / 2D);
        }

        if (alignment == OfficeTextVerticalAlignment.Top) {
            return centerY - (availableHeight / 2D);
        }

        if (alignment == OfficeTextVerticalAlignment.Bottom) {
            return centerY + (availableHeight / 2D) - measured;
        }

        return centerY - (measured / 2D);
    }

    /// <summary>
    /// Rotates a point around the supplied center using clockwise degrees.
    /// </summary>
    /// <param name="point">Point to rotate.</param>
    /// <param name="centerX">Rotation center X coordinate.</param>
    /// <param name="centerY">Rotation center Y coordinate.</param>
    /// <param name="rotationDegrees">Clockwise rotation angle in degrees.</param>
    /// <returns>The rotated point.</returns>
    public static OfficePoint RotatePoint(OfficePoint point, double centerX, double centerY, double rotationDegrees) {
        return OfficeGeometry.RotatePoint(point, centerX, centerY, OfficeGeometry.DegreesToRadians(rotationDegrees));
    }

    private static bool IsFinitePositive(double value) =>
        value > 0D && !double.IsNaN(value) && !double.IsInfinity(value);

    private static double NormalizeNonNegative(double value) =>
        value >= 0D && !double.IsNaN(value) && !double.IsInfinity(value) ? value : 0D;
}
