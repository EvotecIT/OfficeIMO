using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Text;

namespace OfficeIMO.Drawing;

/// <summary>
/// Positioned rich text box inside an <see cref="OfficeDrawing"/> canvas.
/// Coordinates use the drawing's local top-left coordinate space.
/// </summary>
public sealed class OfficeDrawingRichText : OfficeDrawingElement {
    /// <summary>Creates a positioned drawing rich text box.</summary>
    public OfficeDrawingRichText(
        IReadOnlyList<OfficeRichTextRun> runs,
        double x,
        double y,
        double width,
        double height,
        OfficeTextAlignment alignment = OfficeTextAlignment.Left,
        double? lineHeight = null,
        OfficeTextVerticalAlignment verticalAlignment = OfficeTextVerticalAlignment.Top,
        double rotationDegrees = 0D,
        double? rotationCenterX = null,
        double? rotationCenterY = null,
        bool wrapText = true,
        bool shrinkToFit = false,
        bool flipHorizontal = false,
        bool flipVertical = false,
        OfficeTextPadding? padding = null,
        OfficeTextParagraphIndent? paragraphIndent = null) {
        if (runs == null) {
            throw new ArgumentNullException(nameof(runs));
        }

        ValidateFiniteNonNegative(x, nameof(x));
        ValidateFiniteNonNegative(y, nameof(y));
        ValidatePositiveFinite(width, nameof(width));
        ValidatePositiveFinite(height, nameof(height));
        ValidateFinite(rotationDegrees, nameof(rotationDegrees));
        if (lineHeight.HasValue) {
            ValidatePositiveFinite(lineHeight.Value, nameof(lineHeight));
        }

        var copiedRuns = new List<OfficeRichTextRun>(runs.Count);
        var plainText = new StringBuilder();
        for (int i = 0; i < runs.Count; i++) {
            OfficeRichTextRun run = runs[i] ?? throw new ArgumentException("Rich text runs cannot contain null entries.", nameof(runs));
            copiedRuns.Add(new OfficeRichTextRun(run.Text, run.FontSize, run.Color, run.Bold, run.Italic, run.Underline, run.FontFamily, run.Strikethrough, run.BackgroundColor));
            plainText.Append(run.Text);
        }

        Runs = new ReadOnlyCollection<OfficeRichTextRun>(copiedRuns);
        PlainText = plainText.ToString();
        X = x;
        Y = y;
        Width = width;
        Height = height;
        Alignment = alignment;
        LineHeight = lineHeight;
        VerticalAlignment = verticalAlignment;
        RotationDegrees = rotationDegrees;
        RotationCenterX = rotationCenterX ?? x + width / 2D;
        RotationCenterY = rotationCenterY ?? y + height / 2D;
        WrapText = wrapText;
        ShrinkToFit = shrinkToFit;
        FlipHorizontal = flipHorizontal;
        FlipVertical = flipVertical;
        Padding = padding ?? OfficeTextPadding.Empty;
        ParagraphIndent = paragraphIndent ?? OfficeTextParagraphIndent.Empty;
        ValidateFinite(RotationCenterX, nameof(rotationCenterX));
        ValidateFinite(RotationCenterY, nameof(rotationCenterY));
        if (Padding.Horizontal >= Width || Padding.Vertical >= Height) {
            throw new ArgumentOutOfRangeException(nameof(padding), "Rich text padding must leave a positive content rectangle.");
        }
    }

    /// <summary>Styled text runs in paint order.</summary>
    public IReadOnlyList<OfficeRichTextRun> Runs { get; }

    /// <summary>Plain text formed by concatenating all runs.</summary>
    public string PlainText { get; }

    /// <summary>Text box horizontal position inside the drawing.</summary>
    public double X { get; }

    /// <summary>Text box vertical position inside the drawing.</summary>
    public double Y { get; }

    /// <summary>Text box width.</summary>
    public double Width { get; }

    /// <summary>Text box height.</summary>
    public double Height { get; }

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

    /// <summary>Whether renderers should reduce font sizes when text overflows the text box.</summary>
    public bool ShrinkToFit { get; }

    /// <summary>Whether to mirror the rendered rich text horizontally around the frame center.</summary>
    public bool FlipHorizontal { get; }

    /// <summary>Whether to mirror the rendered rich text vertically around the frame center.</summary>
    public bool FlipVertical { get; }

    /// <summary>Insets applied inside the text frame before rich text layout.</summary>
    public OfficeTextPadding Padding { get; }

    /// <summary>First-line and continuation-line offsets applied inside the rich text frame.</summary>
    public OfficeTextParagraphIndent ParagraphIndent { get; }

    /// <summary>Whether the text frame has non-zero padding.</summary>
    public bool HasPadding => !Padding.IsEmpty;

    /// <summary>Whether the text frame has non-zero paragraph indentation.</summary>
    public bool HasParagraphIndent => !ParagraphIndent.IsEmpty;

    /// <summary>Whether the text frame applies rotation or mirroring.</summary>
    public bool HasFrameTransform => Math.Abs(RotationDegrees) > 0.000001D || FlipHorizontal || FlipVertical;

    /// <summary>Creates the reusable frame transform used by SVG and raster renderers.</summary>
    public OfficeImageFrameTransform CreateFrameTransform() => new OfficeImageFrameTransform(RotationDegrees, RotationCenterX, RotationCenterY, FlipHorizontal, FlipVertical);

    /// <summary>Creates a detached copy of this positioned rich text box.</summary>
    public OfficeDrawingRichText Clone() => new OfficeDrawingRichText(Runs, X, Y, Width, Height, Alignment, LineHeight, VerticalAlignment, RotationDegrees, RotationCenterX, RotationCenterY, WrapText, ShrinkToFit, FlipHorizontal, FlipVertical, Padding, ParagraphIndent);

    internal override OfficeDrawingElement CloneElement() => Clone();

    private static void ValidateFiniteNonNegative(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value < 0D) {
            throw new ArgumentOutOfRangeException(paramName, "Drawing rich text coordinates must be finite non-negative numbers.");
        }
    }

    private static void ValidateFinite(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value)) {
            throw new ArgumentOutOfRangeException(paramName, "Drawing rich text rotation values must be finite numbers.");
        }
    }

    private static void ValidatePositiveFinite(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value <= 0D) {
            throw new ArgumentOutOfRangeException(paramName, "Drawing rich text dimensions must be finite positive numbers.");
        }
    }
}
