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
    public OfficeDrawingText(string text, double x, double y, double width, double height, OfficeFontInfo? font = null, OfficeColor? color = null, OfficeTextAlignment alignment = OfficeTextAlignment.Left, double? lineHeight = null, OfficeTextVerticalAlignment verticalAlignment = OfficeTextVerticalAlignment.Top, double rotationDegrees = 0D, double? rotationCenterX = null, double? rotationCenterY = null, bool wrapText = false, bool shrinkToFit = false, bool stackedText = false, bool flipHorizontal = false, bool flipVertical = false, OfficeTextPadding? padding = null, OfficeTextParagraphIndent? paragraphIndent = null) {
        if (text == null) {
            throw new ArgumentNullException(nameof(text));
        }

        ValidateFiniteNonNegative(x, nameof(x));
        ValidateFiniteNonNegative(y, nameof(y));
        ValidatePositiveFinite(width, nameof(width));
        ValidatePositiveFinite(height, nameof(height));
        ValidateFinite(rotationDegrees, nameof(rotationDegrees));
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
        VerticalAlignment = verticalAlignment;
        RotationDegrees = rotationDegrees;
        RotationCenterX = rotationCenterX ?? x + width / 2D;
        RotationCenterY = rotationCenterY ?? y + height / 2D;
        WrapText = wrapText;
        ShrinkToFit = shrinkToFit;
        StackedText = stackedText;
        FlipHorizontal = flipHorizontal;
        FlipVertical = flipVertical;
        Padding = padding ?? OfficeTextPadding.Empty;
        ParagraphIndent = paragraphIndent ?? OfficeTextParagraphIndent.Empty;
        ValidateFinite(RotationCenterX, nameof(rotationCenterX));
        ValidateFinite(RotationCenterY, nameof(rotationCenterY));
        if (Padding.Horizontal >= Width || Padding.Vertical >= Height) {
            throw new ArgumentOutOfRangeException(nameof(padding), "Text padding must leave a positive content rectangle.");
        }
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

    /// <summary>Vertical placement of text inside the drawing text box.</summary>
    public OfficeTextVerticalAlignment VerticalAlignment { get; }

    /// <summary>Clockwise text rotation in degrees.</summary>
    public double RotationDegrees { get; }

    /// <summary>X coordinate of the text rotation center in the drawing's local coordinate space.</summary>
    public double RotationCenterX { get; }

    /// <summary>Y coordinate of the text rotation center in the drawing's local coordinate space.</summary>
    public double RotationCenterY { get; }

    /// <summary>Whether renderers should wrap text to the text box width.</summary>
    public bool WrapText { get; }

    /// <summary>Whether renderers should reduce font size when text overflows the text box.</summary>
    public bool ShrinkToFit { get; }

    /// <summary>Whether renderers should lay text out as upright stacked characters.</summary>
    public bool StackedText { get; }

    /// <summary>Whether renderers should mirror the text frame horizontally around the rotation center.</summary>
    public bool FlipHorizontal { get; }

    /// <summary>Whether renderers should mirror the text frame vertically around the rotation center.</summary>
    public bool FlipVertical { get; }

    /// <summary>Insets applied inside the text frame before text layout.</summary>
    public OfficeTextPadding Padding { get; }

    /// <summary>First-line and continuation-line offsets applied inside the text frame.</summary>
    public OfficeTextParagraphIndent ParagraphIndent { get; }

    /// <summary>Whether the text frame has non-zero padding.</summary>
    public bool HasPadding => !Padding.IsEmpty;

    /// <summary>Whether the text frame has non-zero paragraph indentation.</summary>
    public bool HasParagraphIndent => !ParagraphIndent.IsEmpty;

    /// <summary>Whether the text frame applies rotation or mirroring.</summary>
    public bool HasFrameTransform => Math.Abs(RotationDegrees) > 0.000001D || FlipHorizontal || FlipVertical;

    /// <summary>Creates the reusable destination-space frame transform for this text box.</summary>
    public OfficeImageFrameTransform CreateFrameTransform() => new OfficeImageFrameTransform(RotationDegrees, RotationCenterX, RotationCenterY, FlipHorizontal, FlipVertical);

    /// <summary>Creates a detached copy of this positioned text box.</summary>
    public OfficeDrawingText Clone() => new OfficeDrawingText(Text, X, Y, Width, Height, Font, Color, Alignment, LineHeight, VerticalAlignment, RotationDegrees, RotationCenterX, RotationCenterY, WrapText, ShrinkToFit, StackedText, FlipHorizontal, FlipVertical, Padding, ParagraphIndent);

    internal override OfficeDrawingElement CloneElement() => Clone();

    private static void ValidateFiniteNonNegative(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value < 0D) {
            throw new ArgumentOutOfRangeException(paramName, "Drawing text coordinates must be finite non-negative numbers.");
        }
    }

    private static void ValidateFinite(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value)) {
            throw new ArgumentOutOfRangeException(paramName, "Drawing text rotation values must be finite numbers.");
        }
    }

    private static void ValidatePositiveFinite(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value <= 0D) {
            throw new ArgumentOutOfRangeException(paramName, "Drawing text dimensions must be finite positive numbers.");
        }
    }
}
