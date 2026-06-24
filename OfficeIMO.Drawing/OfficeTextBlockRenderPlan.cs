using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// Shared placement plan for renderers that draw a measured text block inside a center-based rectangle.
/// </summary>
public sealed class OfficeTextBlockRenderPlan {
    private OfficeTextBlockRenderPlan(
        OfficeTextBlockLayout layout,
        double centerX,
        double centerY,
        double width,
        double height,
        OfficeTextAlignment horizontalAlignment,
        OfficeTextVerticalAlignment verticalAlignment) {
        Layout = layout ?? throw new ArgumentNullException(nameof(layout));
        CenterX = centerX;
        CenterY = centerY;
        Width = NormalizePositive(width, Math.Max(layout.Width, 1D));
        Height = NormalizePositive(height, Math.Max(layout.Height, 1D));
        Left = centerX - (Width / 2D);
        Top = centerY - (Height / 2D);
        HorizontalAlignment = horizontalAlignment;
        VerticalAlignment = verticalAlignment;
        TextTop = OfficeTextPlacement.ResolveTop(Top, Height, layout.Height, verticalAlignment);
        AnchorX = OfficeTextPlacement.ResolveAnchorX(Left, Width, horizontalAlignment);
        TextLeft = OfficeTextPlacement.ResolveLeftFromAnchor(AnchorX, layout.Width, horizontalAlignment);
    }

    /// <summary>Measured text layout.</summary>
    public OfficeTextBlockLayout Layout { get; }

    /// <summary>Center X of the available text rectangle.</summary>
    public double CenterX { get; }

    /// <summary>Center Y of the available text rectangle.</summary>
    public double CenterY { get; }

    /// <summary>Left edge of the available text rectangle.</summary>
    public double Left { get; }

    /// <summary>Top edge of the available text rectangle.</summary>
    public double Top { get; }

    /// <summary>Width of the available text rectangle.</summary>
    public double Width { get; }

    /// <summary>Height of the available text rectangle.</summary>
    public double Height { get; }

    /// <summary>Horizontal alignment inside the available rectangle.</summary>
    public OfficeTextAlignment HorizontalAlignment { get; }

    /// <summary>Vertical alignment inside the available rectangle.</summary>
    public OfficeTextVerticalAlignment VerticalAlignment { get; }

    /// <summary>Resolved X anchor used by line renderers.</summary>
    public double AnchorX { get; }

    /// <summary>Resolved left edge of the measured text block.</summary>
    public double TextLeft { get; }

    /// <summary>Resolved top edge of the measured text block.</summary>
    public double TextTop { get; }

    /// <summary>
    /// Creates a placement plan for an already measured layout.
    /// </summary>
    public static OfficeTextBlockRenderPlan CreateFromCenter(
        OfficeTextBlockLayout layout,
        double centerX,
        double centerY,
        double width,
        double height,
        OfficeTextAlignment horizontalAlignment = OfficeTextAlignment.Left,
        OfficeTextVerticalAlignment verticalAlignment = OfficeTextVerticalAlignment.Top) =>
        new(layout, centerX, centerY, width, height, horizontalAlignment, verticalAlignment);

    /// <summary>
    /// Measures wrapped text and creates a placement plan for a center-based rectangle.
    /// </summary>
    public static OfficeTextBlockRenderPlan CreateFittedFromCenter(
        string? text,
        double fontSize,
        double centerX,
        double centerY,
        double width,
        double height,
        Func<string?, double, double> measure,
        OfficeTextAlignment horizontalAlignment = OfficeTextAlignment.Left,
        OfficeTextVerticalAlignment verticalAlignment = OfficeTextVerticalAlignment.Top,
        double lineHeightFactor = 1.2D,
        double minimumFontSize = 5D) {
        if (measure == null) {
            throw new ArgumentNullException(nameof(measure));
        }

        double resolvedWidth = NormalizePositive(width, double.PositiveInfinity);
        double resolvedHeight = NormalizePositive(height, double.PositiveInfinity);
        OfficeTextBlockLayout layout = OfficeTextLayoutEngine.FitWrappedText(
            text,
            fontSize,
            resolvedWidth,
            resolvedHeight,
            lineHeightFactor,
            minimumFontSize,
            measure);
        return CreateFromCenter(layout, centerX, centerY, resolvedWidth, resolvedHeight, horizontalAlignment, verticalAlignment);
    }

    /// <summary>
    /// Creates background bounds around the measured text block.
    /// </summary>
    public OfficeTextBlockBackgroundBounds CreateBackgroundBounds(double horizontalPadding, double verticalPadding) {
        double padX = NormalizeNonNegative(horizontalPadding);
        double padY = NormalizeNonNegative(verticalPadding);
        return new OfficeTextBlockBackgroundBounds(
            TextLeft - padX,
            TextTop - padY,
            Layout.Width + (padX * 2D),
            Layout.Height + (padY * 2D));
    }

    private static double NormalizePositive(double value, double fallback) =>
        value > 0D && !double.IsNaN(value) && !double.IsInfinity(value) ? value : fallback;

    private static double NormalizeNonNegative(double value) =>
        value >= 0D && !double.IsNaN(value) && !double.IsInfinity(value) ? value : 0D;
}

/// <summary>
/// Rectangle around a measured text block, usually used for text background rendering.
/// </summary>
public readonly struct OfficeTextBlockBackgroundBounds {
    /// <summary>Creates text background bounds.</summary>
    public OfficeTextBlockBackgroundBounds(double left, double top, double width, double height) {
        Left = left;
        Top = top;
        Width = width;
        Height = height;
    }

    /// <summary>Left edge of the background rectangle.</summary>
    public double Left { get; }

    /// <summary>Top edge of the background rectangle.</summary>
    public double Top { get; }

    /// <summary>Background rectangle width.</summary>
    public double Width { get; }

    /// <summary>Background rectangle height.</summary>
    public double Height { get; }

    /// <summary>Unrotated rectangle corners in clockwise order.</summary>
    public OfficePoint[] GetCorners() =>
        new[] {
            new OfficePoint(Left, Top),
            new OfficePoint(Left + Width, Top),
            new OfficePoint(Left + Width, Top + Height),
            new OfficePoint(Left, Top + Height)
        };

    /// <summary>
    /// Rotates rectangle corners around a supplied center using clockwise degrees.
    /// </summary>
    public OfficePoint[] GetRotatedCorners(double rotationDegrees, double centerX, double centerY) {
        OfficePoint[] corners = GetCorners();
        if (Math.Abs(rotationDegrees) <= 0.000001D) {
            return corners;
        }

        for (int i = 0; i < corners.Length; i++) {
            corners[i] = OfficeTextPlacement.RotatePoint(corners[i], centerX, centerY, rotationDegrees);
        }

        return corners;
    }
}
