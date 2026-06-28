using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// Shared placement plan for renderers that draw a measured text block inside a center-based rectangle.
/// </summary>
public sealed class OfficeTextBlockRenderPlan {
    private OfficeTextBlockRenderPlan(
        OfficeTextBlockLayout layout,
        double left,
        double top,
        double width,
        double height,
        OfficeTextAlignment horizontalAlignment,
        OfficeTextVerticalAlignment verticalAlignment) {
        Layout = layout ?? throw new ArgumentNullException(nameof(layout));
        Width = NormalizePositive(width, Math.Max(layout.Width, 1D));
        Height = NormalizePositive(height, Math.Max(layout.Height, 1D));
        Left = NormalizeCoordinate(left);
        Top = NormalizeCoordinate(top);
        CenterX = Left + (Width / 2D);
        CenterY = Top + (Height / 2D);
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
        CreateFromRectangle(
            layout,
            centerX - (NormalizePositive(width, Math.Max(layout?.Width ?? 1D, 1D)) / 2D),
            centerY - (NormalizePositive(height, Math.Max(layout?.Height ?? 1D, 1D)) / 2D),
            width,
            height,
            horizontalAlignment,
            verticalAlignment);

    /// <summary>
    /// Creates a placement plan for an already measured layout inside a left/top-based rectangle.
    /// </summary>
    public static OfficeTextBlockRenderPlan CreateFromRectangle(
        OfficeTextBlockLayout layout,
        double left,
        double top,
        double width,
        double height,
        OfficeTextAlignment horizontalAlignment = OfficeTextAlignment.Left,
        OfficeTextVerticalAlignment verticalAlignment = OfficeTextVerticalAlignment.Top) =>
        new(layout, left, top, width, height, horizontalAlignment, verticalAlignment);

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
    /// Measures a text block and creates a placement plan for a left/top-based rectangle.
    /// </summary>
    /// <param name="text">Text to lay out.</param>
    /// <param name="fontSize">Initial font size passed to <paramref name="measure"/>.</param>
    /// <param name="left">Left edge of the available text rectangle.</param>
    /// <param name="top">Top edge of the available text rectangle.</param>
    /// <param name="width">Available text rectangle width.</param>
    /// <param name="height">Available text rectangle height.</param>
    /// <param name="measure">Measurement delegate matching <see cref="OfficeRasterCanvas.MeasureText(string?, double)"/>.</param>
    /// <param name="horizontalAlignment">Horizontal alignment inside the rectangle.</param>
    /// <param name="verticalAlignment">Vertical alignment inside the rectangle.</param>
    /// <param name="lineHeightFactor">Multiplier used to derive line height from font size.</param>
    /// <param name="minimumFontSize">Minimum font size when shrink-to-fit is enabled.</param>
    /// <param name="wrap">Whether soft wrapping is enabled.</param>
    /// <param name="forceSingleLine">Whether line breaks should be normalized to spaces and wrapping disabled.</param>
    /// <param name="shrinkToFit">Whether single-line text should reduce font size to fit the requested width.</param>
    public static OfficeTextBlockRenderPlan CreateTextBlockFromRectangle(
        string? text,
        double fontSize,
        double left,
        double top,
        double width,
        double height,
        Func<string?, double, double> measure,
        OfficeTextAlignment horizontalAlignment = OfficeTextAlignment.Left,
        OfficeTextVerticalAlignment verticalAlignment = OfficeTextVerticalAlignment.Top,
        double lineHeightFactor = 1.2D,
        double minimumFontSize = 5D,
        bool wrap = false,
        bool forceSingleLine = false,
        bool shrinkToFit = false) =>
        CreateTextBlockFromRectangle(
            text,
            fontSize,
            left,
            top,
            width,
            height,
            measure,
            horizontalAlignment,
            verticalAlignment,
            lineHeightFactor,
            minimumFontSize,
            wrap,
            forceSingleLine,
            shrinkToFit,
            OfficeTextOverflowBehavior.Ellipsis);

    /// <summary>
    /// Measures a text block and creates a placement plan for a left/top-based rectangle.
    /// </summary>
    /// <param name="text">Text to lay out.</param>
    /// <param name="fontSize">Initial font size passed to <paramref name="measure"/>.</param>
    /// <param name="left">Left edge of the available text rectangle.</param>
    /// <param name="top">Top edge of the available text rectangle.</param>
    /// <param name="width">Available text rectangle width.</param>
    /// <param name="height">Available text rectangle height.</param>
    /// <param name="measure">Measurement delegate matching <see cref="OfficeRasterCanvas.MeasureText(string?, double)"/>.</param>
    /// <param name="horizontalAlignment">Horizontal alignment inside the rectangle.</param>
    /// <param name="verticalAlignment">Vertical alignment inside the rectangle.</param>
    /// <param name="lineHeightFactor">Multiplier used to derive line height from font size.</param>
    /// <param name="minimumFontSize">Minimum font size when shrink-to-fit is enabled.</param>
    /// <param name="wrap">Whether soft wrapping is enabled.</param>
    /// <param name="forceSingleLine">Whether line breaks should be normalized to spaces and wrapping disabled.</param>
    /// <param name="shrinkToFit">Whether single-line text should reduce font size to fit the requested width.</param>
    /// <param name="overflowBehavior">How overflowing text should be represented in the returned layout.</param>
    public static OfficeTextBlockRenderPlan CreateTextBlockFromRectangle(
        string? text,
        double fontSize,
        double left,
        double top,
        double width,
        double height,
        Func<string?, double, double> measure,
        OfficeTextAlignment horizontalAlignment,
        OfficeTextVerticalAlignment verticalAlignment,
        double lineHeightFactor,
        double minimumFontSize,
        bool wrap,
        bool forceSingleLine,
        bool shrinkToFit,
        OfficeTextOverflowBehavior overflowBehavior) {
        if (measure == null) {
            throw new ArgumentNullException(nameof(measure));
        }

        double resolvedWidth = NormalizePositive(width, 1D);
        double resolvedHeight = NormalizePositive(height, 1D);
        OfficeTextBlockLayout layout = OfficeTextLayoutEngine.LayoutTextBlock(
            text,
            fontSize,
            resolvedWidth,
            resolvedHeight,
            lineHeightFactor,
            minimumFontSize,
            measure,
            wrap,
            forceSingleLine,
            shrinkToFit,
            overflowBehavior);
        return CreateFromRectangle(layout, left, top, resolvedWidth, resolvedHeight, horizontalAlignment, verticalAlignment);
    }

    /// <summary>
    /// Measures stacked text and creates a placement plan for a left/top-based rectangle.
    /// </summary>
    public static OfficeTextBlockRenderPlan CreateStackedTextBlockFromRectangle(
        string? text,
        double fontSize,
        double left,
        double top,
        double width,
        double height,
        Func<string?, double, double> measure,
        OfficeTextAlignment horizontalAlignment = OfficeTextAlignment.Center,
        OfficeTextVerticalAlignment verticalAlignment = OfficeTextVerticalAlignment.Top,
        double lineHeightFactor = 1.2D,
        double minimumFontSize = 5D,
        bool shrinkToFit = false) {
        if (measure == null) {
            throw new ArgumentNullException(nameof(measure));
        }

        double resolvedWidth = NormalizePositive(width, 1D);
        double resolvedHeight = NormalizePositive(height, 1D);
        OfficeTextBlockLayout layout = OfficeTextLayoutEngine.LayoutStackedTextBlock(
            text,
            fontSize,
            resolvedWidth,
            resolvedHeight,
            lineHeightFactor,
            minimumFontSize,
            measure,
            shrinkToFit);
        return CreateFromRectangle(layout, left, top, resolvedWidth, resolvedHeight, horizontalAlignment, verticalAlignment);
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

    private static double NormalizeCoordinate(double value) =>
        !double.IsNaN(value) && !double.IsInfinity(value) ? value : 0D;
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
