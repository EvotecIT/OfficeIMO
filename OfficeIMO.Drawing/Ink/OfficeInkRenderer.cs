using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>Settings used to project reusable ink into an <see cref="OfficeDrawing"/>.</summary>
public sealed class OfficeInkRenderOptions {
    /// <summary>Default opacity multiplier used for strokes marked as highlighters.</summary>
    public const double DefaultHighlighterOpacityFactor = 0.4D;

    /// <summary>Whether normalized pressure changes segment thickness.</summary>
    public bool UsePressure { get; set; } = true;

    /// <summary>Smallest pressure multiplier applied when pressure is present.</summary>
    public double MinimumPressureFactor { get; set; } = 0.25D;

    /// <summary>Opacity multiplier applied to strokes marked as highlighters.</summary>
    public double HighlighterOpacityFactor { get; set; } = DefaultHighlighterOpacityFactor;

    /// <summary>Creates a detached copy.</summary>
    public OfficeInkRenderOptions Clone() => new OfficeInkRenderOptions {
        UsePressure = UsePressure,
        MinimumPressureFactor = MinimumPressureFactor,
        HighlighterOpacityFactor = HighlighterOpacityFactor
    };

    internal void Validate() {
        ValidateUnitInterval(MinimumPressureFactor, nameof(MinimumPressureFactor));
        ValidateUnitInterval(HighlighterOpacityFactor, nameof(HighlighterOpacityFactor));
    }

    private static void ValidateUnitInterval(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value < 0D || value > 1D) {
            throw new ArgumentOutOfRangeException(paramName, "Ink rendering factors must be from 0 through 1.");
        }
    }
}

/// <summary>Projects format-neutral ink strokes into the shared drawing scene.</summary>
public static class OfficeInkRenderer {
    /// <summary>Returns the effective rendered opacity, including color alpha and highlighter behavior.</summary>
    public static double GetEffectiveOpacity(
        OfficeInkStroke stroke,
        double highlighterOpacityFactor = OfficeInkRenderOptions.DefaultHighlighterOpacityFactor) {
        if (stroke == null) throw new ArgumentNullException(nameof(stroke));
        stroke.ValidateStyle();
        if (double.IsNaN(highlighterOpacityFactor) || double.IsInfinity(highlighterOpacityFactor) ||
            highlighterOpacityFactor < 0D || highlighterOpacityFactor > 1D) {
            throw new ArgumentOutOfRangeException(nameof(highlighterOpacityFactor), "The highlighter opacity factor must be from 0 through 1.");
        }
        double highlighterFactor = stroke.IsHighlighter ? highlighterOpacityFactor : 1D;
        return stroke.Opacity * highlighterFactor * stroke.Color.A / byte.MaxValue;
    }

    /// <summary>Creates a new drawing containing the supplied ink.</summary>
    public static OfficeDrawing Render(
        OfficeInkDocument ink,
        double width,
        double height,
        OfficeInkRenderOptions? options = null) {
        var drawing = new OfficeDrawing(width, height);
        AddToDrawing(drawing, ink, options: options);
        return drawing;
    }

    /// <summary>Adds ink to an existing drawing at an optional destination offset.</summary>
    public static void AddToDrawing(
        OfficeDrawing drawing,
        OfficeInkDocument ink,
        double offsetX = 0D,
        double offsetY = 0D,
        OfficeInkRenderOptions? options = null) {
        if (drawing == null) throw new ArgumentNullException(nameof(drawing));
        if (ink == null) throw new ArgumentNullException(nameof(ink));
        ValidateFinite(offsetX, nameof(offsetX));
        ValidateFinite(offsetY, nameof(offsetY));
        OfficeInkRenderOptions effective = options?.Clone() ?? new OfficeInkRenderOptions();
        effective.Validate();

        for (int index = 0; index < ink.Strokes.Count; index++) {
            AddStroke(drawing, ink.Strokes[index], offsetX, offsetY, effective);
        }
    }

    private static void AddStroke(
        OfficeDrawing drawing,
        OfficeInkStroke stroke,
        double offsetX,
        double offsetY,
        OfficeInkRenderOptions options) {
        stroke.ValidateStyle();
        if (stroke.Points.Count == 0 || stroke.Opacity <= 0D || stroke.Color.A == 0) return;

        OfficeTransform transform = stroke.Transform ?? OfficeTransform.Identity;
        (double transformedTipWidth, double transformedTipHeight) = stroke.GetTransformedTipDimensions();
        var points = new List<OfficeInkPoint>(stroke.Points.Count);
        for (int index = 0; index < stroke.Points.Count; index++) {
            OfficeInkPoint transformed = stroke.Points[index].Transform(transform);
            points.Add(new OfficeInkPoint(
                transformed.X + offsetX,
                transformed.Y + offsetY,
                transformed.Pressure,
                transformed.TiltX,
                transformed.TiltY,
                transformed.Timestamp));
        }

        double opacity = GetEffectiveOpacity(stroke, options.HighlighterOpacityFactor);
        OfficeColor renderColor = OfficeColor.FromRgb(stroke.Color.R, stroke.Color.G, stroke.Color.B);
        if (points.Count == 1) {
            AddDot(drawing, stroke, renderColor, points[0], transformedTipWidth, transformedTipHeight, opacity, options);
            return;
        }

        for (int index = 1; index < points.Count; index++) {
            OfficeInkPoint from = points[index - 1];
            OfficeInkPoint to = points[index];
            double x1 = from.X;
            double y1 = from.Y;
            double x2 = to.X;
            double y2 = to.Y;
            if (x1.Equals(x2) && y1.Equals(y2)) continue;
            double pressure = options.UsePressure && !stroke.IgnorePressure
                ? ResolvePressure(from.Pressure, to.Pressure, options.MinimumPressureFactor)
                : 1D;
            double deltaX = to.X - from.X;
            double deltaY = to.Y - from.Y;
            double thickness = Math.Max(0.01D, stroke.GetTransformedTipExtent(-deltaY, deltaX) * pressure);
            double clipHalfWidth = transformedTipWidth * pressure / 2D;
            double clipHalfHeight = transformedTipHeight * pressure / 2D;
            if (!TryClipLine(-clipHalfWidth, -clipHalfHeight, drawing.Width + clipHalfWidth, drawing.Height + clipHalfHeight, ref x1, ref y1, ref x2, ref y2)) continue;
            if (x1.Equals(x2) && y1.Equals(y2)) continue;

            OfficeShape shape = OfficeShape.Line(x1, y1, x2, y2);
            shape.StrokeColor = renderColor;
            shape.StrokeWidth = thickness;
            shape.StrokeOpacity = opacity;
            shape.StrokeLineCap = stroke.TipShape == OfficeInkTipShape.Rectangle
                ? OfficeStrokeLineCap.Square
                : OfficeStrokeLineCap.Round;
            shape.StrokeLineJoin = OfficeStrokeLineJoin.Round;
            bool exceedsCanvas = Math.Min(x1, x2) < 0D || Math.Min(y1, y2) < 0D ||
                Math.Max(x1, x2) > drawing.Width || Math.Max(y1, y2) > drawing.Height;
            if (exceedsCanvas) drawing.AddShapeForClippedRendering(shape, Math.Min(x1, x2), Math.Min(y1, y2));
            else drawing.AddShape(shape, Math.Min(x1, x2), Math.Min(y1, y2));
        }
    }

    private static void AddDot(
        OfficeDrawing drawing,
        OfficeInkStroke stroke,
        OfficeColor renderColor,
        OfficeInkPoint point,
        double transformedTipWidth,
        double transformedTipHeight,
        double opacity,
        OfficeInkRenderOptions options) {
        double pressure = options.UsePressure && !stroke.IgnorePressure
            ? ResolvePressure(point.Pressure, point.Pressure, options.MinimumPressureFactor)
            : 1D;
        double width = Math.Max(0.01D, transformedTipWidth * pressure);
        double height = Math.Max(0.01D, transformedTipHeight * pressure);
        double x = point.X - width / 2D;
        double y = point.Y - height / 2D;
        if (x >= drawing.Width || y >= drawing.Height || x + width <= 0D || y + height <= 0D) return;

        OfficeShape shape = stroke.TipShape == OfficeInkTipShape.Rectangle
            ? OfficeShape.Rectangle(width, height)
            : OfficeShape.Ellipse(width, height);
        shape.FillColor = renderColor;
        shape.FillOpacity = opacity;
        if (x < 0D || y < 0D || x + width > drawing.Width || y + height > drawing.Height) {
            drawing.AddShapeForClippedRendering(shape, x, y);
        } else {
            drawing.AddShape(shape, x, y);
        }
    }

    private static double ResolvePressure(double? first, double? second, double minimum) {
        if (!first.HasValue && !second.HasValue) return 1D;
        double value = first.HasValue && second.HasValue
            ? (first.Value + second.Value) / 2D
            : first ?? second ?? 1D;
        value = Math.Max(0D, Math.Min(1D, value));
        return minimum + (1D - minimum) * value;
    }

    private static bool TryClipLine(
        double left,
        double top,
        double right,
        double bottom,
        ref double x1,
        ref double y1,
        ref double x2,
        ref double y2) {
        int code1 = RegionCode(left, top, right, bottom, x1, y1);
        int code2 = RegionCode(left, top, right, bottom, x2, y2);
        while (true) {
            if ((code1 | code2) == 0) return true;
            if ((code1 & code2) != 0) return false;

            int outside = code1 != 0 ? code1 : code2;
            double x;
            double y;
            if ((outside & 8) != 0) {
                x = x1 + (x2 - x1) * (bottom - y1) / (y2 - y1);
                y = bottom;
            } else if ((outside & 4) != 0) {
                x = x1 + (x2 - x1) * (top - y1) / (y2 - y1);
                y = top;
            } else if ((outside & 2) != 0) {
                y = y1 + (y2 - y1) * (right - x1) / (x2 - x1);
                x = right;
            } else {
                y = y1 + (y2 - y1) * (left - x1) / (x2 - x1);
                x = left;
            }

            if (outside == code1) {
                x1 = x;
                y1 = y;
                code1 = RegionCode(left, top, right, bottom, x1, y1);
            } else {
                x2 = x;
                y2 = y;
                code2 = RegionCode(left, top, right, bottom, x2, y2);
            }
        }
    }

    private static int RegionCode(double left, double top, double right, double bottom, double x, double y) {
        int code = 0;
        if (x < left) code |= 1;
        else if (x > right) code |= 2;
        if (y < top) code |= 4;
        else if (y > bottom) code |= 8;
        return code;
    }

    private static void ValidateFinite(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value)) {
            throw new ArgumentOutOfRangeException(paramName, "Ink offsets must be finite numbers.");
        }
    }
}
