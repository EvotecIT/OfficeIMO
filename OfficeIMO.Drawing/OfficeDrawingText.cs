using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// Positioned text box inside an <see cref="OfficeDrawing"/> canvas.
/// Coordinates use the drawing's local top-left coordinate space.
/// </summary>
public sealed class OfficeDrawingText : OfficeDrawingElement {
    /// <summary>
    /// Creates a positioned drawing text box.
    /// </summary>
    public OfficeDrawingText(string text, double x, double y, double width, double height, OfficeFontInfo? font = null, OfficeColor? color = null, OfficeTextAlignment alignment = OfficeTextAlignment.Left, double? lineHeight = null) {
        if (text == null) {
            throw new ArgumentNullException(nameof(text));
        }

        ValidateFiniteNonNegative(x, nameof(x));
        ValidateFiniteNonNegative(y, nameof(y));
        ValidatePositiveFinite(width, nameof(width));
        ValidatePositiveFinite(height, nameof(height));
        if (lineHeight.HasValue) {
            ValidatePositiveFinite(lineHeight.Value, nameof(lineHeight));
        }

        Text = text;
        X = x;
        Y = y;
        Width = width;
        Height = height;
        Font = font ?? OfficeFontInfo.Default;
        Color = color;
        Alignment = alignment;
        LineHeight = lineHeight;
    }

    /// <summary>Text content.</summary>
    public string Text { get; }

    /// <summary>Text box horizontal position inside the drawing.</summary>
    public double X { get; }

    /// <summary>Text box vertical position inside the drawing.</summary>
    public double Y { get; }

    /// <summary>Text box width.</summary>
    public double Width { get; }

    /// <summary>Text box height.</summary>
    public double Height { get; }

    /// <summary>Font descriptor for the text.</summary>
    public OfficeFontInfo Font { get; }

    /// <summary>Optional text color. Null lets renderers choose their default text color.</summary>
    public OfficeColor? Color { get; }

    /// <summary>Horizontal text alignment inside the text box.</summary>
    public OfficeTextAlignment Alignment { get; }

    /// <summary>Optional line height in drawing units. Null lets renderers use their default line height.</summary>
    public double? LineHeight { get; }

    /// <summary>Creates a detached copy of this positioned text box.</summary>
    public OfficeDrawingText Clone() => new OfficeDrawingText(Text, X, Y, Width, Height, Font, Color, Alignment, LineHeight);

    internal override OfficeDrawingElement CloneElement() => Clone();

    private static void ValidateFiniteNonNegative(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value < 0D) {
            throw new ArgumentOutOfRangeException(paramName, "Drawing text coordinates must be finite non-negative numbers.");
        }
    }

    private static void ValidatePositiveFinite(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value <= 0D) {
            throw new ArgumentOutOfRangeException(paramName, "Drawing text dimensions must be finite positive numbers.");
        }
    }
}
