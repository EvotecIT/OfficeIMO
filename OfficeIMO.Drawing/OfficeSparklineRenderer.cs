using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeIMO.Drawing;

/// <summary>
/// Shared dependency-free sparkline renderer for raster and SVG output.
/// </summary>
public static class OfficeSparklineRenderer {
    /// <summary>
    /// Draws a sparkline on a raster canvas.
    /// </summary>
    public static void DrawRaster(
        OfficeRasterCanvas canvas,
        double x,
        double y,
        double width,
        double height,
        IReadOnlyList<double> values,
        OfficeSparklineKind kind,
        OfficeSparklineStyle? style = null) {
        if (canvas == null) {
            throw new ArgumentNullException(nameof(canvas));
        }

        if (values == null) {
            throw new ArgumentNullException(nameof(values));
        }

        if (values.Count == 0 || width <= 0D || height <= 0D) {
            return;
        }

        OfficeSparklineStyle resolvedStyle = style ?? new OfficeSparklineStyle();
        SparklineBounds bounds = ResolveBounds(x, y, width, height, resolvedStyle);
        ValueRange range = ResolveRange(values);

        if (kind == OfficeSparklineKind.Column || kind == OfficeSparklineKind.WinLoss) {
            DrawRasterColumns(canvas, bounds, values, range, kind == OfficeSparklineKind.WinLoss, resolvedStyle);
            return;
        }

        DrawRasterLine(canvas, bounds, values, range, resolvedStyle);
    }

    /// <summary>
    /// Appends SVG elements for a sparkline.
    /// </summary>
    public static StringBuilder AppendSvg(
        StringBuilder builder,
        double x,
        double y,
        double width,
        double height,
        IReadOnlyList<double> values,
        OfficeSparklineKind kind,
        OfficeSparklineStyle? style = null) {
        if (builder == null) {
            throw new ArgumentNullException(nameof(builder));
        }

        if (values == null) {
            throw new ArgumentNullException(nameof(values));
        }

        if (values.Count == 0 || width <= 0D || height <= 0D) {
            return builder;
        }

        OfficeSparklineStyle resolvedStyle = style ?? new OfficeSparklineStyle();
        SparklineBounds bounds = ResolveBounds(x, y, width, height, resolvedStyle);
        ValueRange range = ResolveRange(values);

        if (kind == OfficeSparklineKind.Column || kind == OfficeSparklineKind.WinLoss) {
            AppendSvgColumns(builder, bounds, values, range, kind == OfficeSparklineKind.WinLoss, resolvedStyle);
            return builder;
        }

        AppendSvgLine(builder, bounds, values, range, resolvedStyle);
        return builder;
    }

    private static void DrawRasterLine(OfficeRasterCanvas canvas, SparklineBounds bounds, IReadOnlyList<double> values, ValueRange range, OfficeSparklineStyle style) {
        List<OfficePoint> points = BuildLinePoints(bounds, values, range);
        DrawRasterAxis(canvas, bounds, range, style);

        if (points.Count == 1) {
            DrawRasterMarker(canvas, points[0], ResolvePointStyle(style, 0), Math.Max(3D, style.MarkerDiameter));
            return;
        }

        canvas.DrawPolyline(points, style.SeriesColor, Math.Max(1D, style.LineStrokeWidth));
        for (int i = 0; i < points.Count; i++) {
            OfficeSparklinePointStyle pointStyle = ResolvePointStyle(style, i);
            if (pointStyle.ShowMarker) {
                DrawRasterMarker(canvas, points[i], pointStyle, Math.Max(3D, style.MarkerDiameter));
            }
        }
    }

    private static void DrawRasterColumns(OfficeRasterCanvas canvas, SparklineBounds bounds, IReadOnlyList<double> values, ValueRange range, bool winLoss, OfficeSparklineStyle style) {
        DrawRasterAxis(canvas, bounds, range, style);
        double baseline = ResolveZeroY(range, bounds.Top, bounds.Height);
        double slot = bounds.Width / values.Count;
        double barWidth = Math.Max(1D, slot * ClampPositiveRatio(style.ColumnWidthRatio, 0.62D));
        double valueExtent = Math.Max(Math.Abs(range.Max), Math.Abs(range.Min));
        if (valueExtent < 0.000001D) {
            valueExtent = 1D;
        }

        for (int i = 0; i < values.Count; i++) {
            double value = values[i];
            double barHeight = winLoss
                ? Math.Max(1D, bounds.Height * ClampPositiveRatio(style.WinLossHeightRatio, 0.42D))
                : Math.Max(1D, Math.Abs(value) / valueExtent * bounds.Height);
            double barX = bounds.Left + (slot * i) + ((slot - barWidth) / 2D);
            double barY = value < 0D ? baseline : baseline - barHeight;
            canvas.FillRectangle(barX, barY, barWidth, barHeight, ResolvePointStyle(style, i).Color);
        }
    }

    private static void DrawRasterAxis(OfficeRasterCanvas canvas, SparklineBounds bounds, ValueRange range, OfficeSparklineStyle style) {
        if (!ShouldDrawAxis(range, style)) {
            return;
        }

        double y = ResolveZeroY(range, bounds.Top, bounds.Height);
        canvas.DrawLine(bounds.AxisLeft, y, bounds.AxisRight, y, style.AxisColor, Math.Max(1D, style.AxisStrokeWidth));
    }

    private static void DrawRasterMarker(OfficeRasterCanvas canvas, OfficePoint point, OfficeSparklinePointStyle style, double markerDiameter) {
        canvas.FillEllipse(point.X - (markerDiameter / 2D), point.Y - (markerDiameter / 2D), markerDiameter, markerDiameter, style.Color);
    }

    private static void AppendSvgLine(StringBuilder builder, SparklineBounds bounds, IReadOnlyList<double> values, ValueRange range, OfficeSparklineStyle style) {
        List<OfficePoint> points = BuildLinePoints(bounds, values, range);
        AppendSvgAxis(builder, bounds, range, style);

        if (points.Count > 1) {
            var attributes = new StringBuilder();
            attributes
                .AppendAttribute("fill", "none")
                .AppendPaintAttribute("stroke", style.SeriesColor)
                .AppendNumberAttribute("stroke-width", Math.Max(1D, style.LineStrokeWidth));
            builder.AppendPolylineElement(points, attributes.ToString());
        }

        for (int i = 0; i < points.Count; i++) {
            OfficeSparklinePointStyle pointStyle = ResolvePointStyle(style, i);
            if (points.Count == 1 || pointStyle.ShowMarker) {
                AppendSvgMarker(builder, points[i], pointStyle, Math.Max(3D, style.MarkerDiameter) / 2D);
            }
        }
    }

    private static void AppendSvgColumns(StringBuilder builder, SparklineBounds bounds, IReadOnlyList<double> values, ValueRange range, bool winLoss, OfficeSparklineStyle style) {
        AppendSvgAxis(builder, bounds, range, style);
        double baseline = ResolveZeroY(range, bounds.Top, bounds.Height);
        double slot = bounds.Width / values.Count;
        double barWidth = Math.Max(1D, slot * ClampPositiveRatio(style.ColumnWidthRatio, 0.62D));
        double valueExtent = Math.Max(Math.Abs(range.Max), Math.Abs(range.Min));
        if (valueExtent < 0.000001D) {
            valueExtent = 1D;
        }

        for (int i = 0; i < values.Count; i++) {
            double value = values[i];
            double barHeight = winLoss
                ? Math.Max(1D, bounds.Height * ClampPositiveRatio(style.WinLossHeightRatio, 0.42D))
                : Math.Max(1D, Math.Abs(value) / valueExtent * bounds.Height);
            double barX = bounds.Left + (slot * i) + ((slot - barWidth) / 2D);
            double barY = value < 0D ? baseline : baseline - barHeight;
            var attributes = new StringBuilder();
            attributes.AppendPaintAttribute("fill", ResolvePointStyle(style, i).Color);
            builder.AppendRectElement(barX, barY, barWidth, barHeight, attributes.ToString());
        }
    }

    private static void AppendSvgAxis(StringBuilder builder, SparklineBounds bounds, ValueRange range, OfficeSparklineStyle style) {
        if (!ShouldDrawAxis(range, style)) {
            return;
        }

        double y = ResolveZeroY(range, bounds.Top, bounds.Height);
        builder.AppendLineElement(bounds.AxisLeft, y, bounds.AxisRight, y, style.AxisColor, Math.Max(1D, style.AxisStrokeWidth));
    }

    private static void AppendSvgMarker(StringBuilder builder, OfficePoint point, OfficeSparklinePointStyle style, double radius) {
        builder.AppendCircleElement(point.X, point.Y, radius, style.Color);
    }

    private static List<OfficePoint> BuildLinePoints(SparklineBounds bounds, IReadOnlyList<double> values, ValueRange range) {
        double valueRange = Math.Abs(range.Max - range.Min) < 0.000001D ? 1D : range.Max - range.Min;
        var points = new List<OfficePoint>(values.Count);
        for (int i = 0; i < values.Count; i++) {
            double x = values.Count == 1 ? bounds.Left + (bounds.Width / 2D) : bounds.Left + (bounds.Width * i / (values.Count - 1D));
            double y = bounds.Top + bounds.Height - (((values[i] - range.Min) / valueRange) * bounds.Height);
            points.Add(new OfficePoint(x, y));
        }

        return points;
    }

    private static ValueRange ResolveRange(IReadOnlyList<double> values) {
        double min = values[0];
        double max = values[0];
        for (int i = 1; i < values.Count; i++) {
            min = Math.Min(min, values[i]);
            max = Math.Max(max, values[i]);
        }

        return new ValueRange(min, max);
    }

    private static SparklineBounds ResolveBounds(double x, double y, double width, double height, OfficeSparklineStyle style) {
        double padding = Math.Max(0D, style.Padding);
        double contentWidth = Math.Max(1D, width - (padding * 2D));
        double contentHeight = Math.Max(1D, height - (padding * 2D));
        double axisInset = Math.Max(0D, style.AxisInset);
        return new SparklineBounds(
            x + padding,
            y + padding,
            contentWidth,
            contentHeight,
            x + axisInset,
            x + width - axisInset);
    }

    private static OfficeSparklinePointStyle ResolvePointStyle(OfficeSparklineStyle style, int index) {
        if (style.PointStyles != null && index >= 0 && index < style.PointStyles.Count) {
            return style.PointStyles[index];
        }

        return new OfficeSparklinePointStyle(style.SeriesColor);
    }

    private static bool ShouldDrawAxis(ValueRange range, OfficeSparklineStyle style) =>
        style.DisplayAxis && range.Min < 0D && range.Max > 0D;

    private static double ResolveZeroY(ValueRange range, double top, double height) {
        if (range.Min >= 0D) {
            return top + height;
        }

        if (range.Max <= 0D) {
            return top;
        }

        return top + height - ((0D - range.Min) / (range.Max - range.Min) * height);
    }

    private static double ClampPositiveRatio(double value, double fallback) =>
        double.IsNaN(value) || double.IsInfinity(value) || value <= 0D ? fallback : value;

    private readonly struct SparklineBounds {
        internal SparklineBounds(double left, double top, double width, double height, double axisLeft, double axisRight) {
            Left = left;
            Top = top;
            Width = width;
            Height = height;
            AxisLeft = axisLeft;
            AxisRight = axisRight;
        }

        internal double Left { get; }

        internal double Top { get; }

        internal double Width { get; }

        internal double Height { get; }

        internal double AxisLeft { get; }

        internal double AxisRight { get; }
    }

    private readonly struct ValueRange {
        internal ValueRange(double min, double max) {
            Min = min;
            Max = max;
        }

        internal double Min { get; }

        internal double Max { get; }
    }
}
