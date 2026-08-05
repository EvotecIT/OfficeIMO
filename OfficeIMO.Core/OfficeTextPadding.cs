using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// Text insets applied inside an Office drawing text frame before layout.
/// </summary>
public readonly struct OfficeTextPadding {
    /// <summary>Creates text-frame padding in drawing units.</summary>
    public OfficeTextPadding(double left, double top, double right, double bottom) {
        ValidateNonNegativeFinite(left, nameof(left));
        ValidateNonNegativeFinite(top, nameof(top));
        ValidateNonNegativeFinite(right, nameof(right));
        ValidateNonNegativeFinite(bottom, nameof(bottom));

        Left = left;
        Top = top;
        Right = right;
        Bottom = bottom;
    }

    /// <summary>No text padding.</summary>
    public static OfficeTextPadding Empty { get; } = new OfficeTextPadding(0D, 0D, 0D, 0D);

    /// <summary>Left inset.</summary>
    public double Left { get; }

    /// <summary>Top inset.</summary>
    public double Top { get; }

    /// <summary>Right inset.</summary>
    public double Right { get; }

    /// <summary>Bottom inset.</summary>
    public double Bottom { get; }

    /// <summary>Total horizontal inset.</summary>
    public double Horizontal => Left + Right;

    /// <summary>Total vertical inset.</summary>
    public double Vertical => Top + Bottom;

    /// <summary>Whether all insets are zero.</summary>
    public bool IsEmpty => Left == 0D && Top == 0D && Right == 0D && Bottom == 0D;

    /// <summary>Returns the same padding scaled by a rendering factor.</summary>
    public OfficeTextPadding Scale(double scale) {
        double factor = scale > 0D && !double.IsNaN(scale) && !double.IsInfinity(scale) ? scale : 1D;
        return new OfficeTextPadding(Left * factor, Top * factor, Right * factor, Bottom * factor);
    }

    private static void ValidateNonNegativeFinite(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value < 0D) {
            throw new ArgumentOutOfRangeException(paramName, "Text padding values must be finite non-negative numbers.");
        }
    }
}
