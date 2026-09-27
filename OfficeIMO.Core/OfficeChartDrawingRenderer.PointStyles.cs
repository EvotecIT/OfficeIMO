using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeChartDrawingRenderer {
    private const string AreaPointStyleWarning = "Per-point area styles are not applied by the shared renderer; the series fill remains in use.";

    private static bool HasUnsupportedAreaPointStyles(OfficeChartSnapshot snapshot) => snapshot.Data.Series.Any(series =>
        IsAreaChart(series.RenderKind ?? snapshot.ChartKind) && series.PointStyles?.Any(style => style != null &&
            (style.NoFill || style.FillColor.HasValue || style.Hatch.HasValue || style.OutlineColor.HasValue ||
                style.OutlineWidth.HasValue || style.ShowOutline.HasValue)) == true);

    private static OfficeChartPointStyle? GetPointStyle(OfficeChartSeries series, int index) =>
        GetPointStyle(series.PointStyles, index);

    private static OfficeChartPointStyle? GetPointStyle(IReadOnlyList<OfficeChartPointStyle?>? styles, int index) =>
        styles != null && index >= 0 && index < styles.Count ? styles[index] : null;

    private static void AddStyledPointShape(OfficeDrawing drawing, OfficeShape shape, double x, double y,
        OfficeColor color, OfficeChartPointStyle? style, OfficeColor? defaultOutline, double defaultWidth) {
        OfficeColor? fill = style?.NoFill == true ? null : style?.FillColor ?? color;
        OfficeColor? outline = style?.ShowOutline == false ? null : style?.OutlineColor ?? defaultOutline;
        if (style?.ShowOutline == true && !outline.HasValue) outline = OfficeColor.Black;
        double width = style?.OutlineWidth ?? defaultWidth;
        AddShape(drawing, shape.Clone(), x, y, fill, style?.Hatch == null ? outline : null, width);
        if (style?.Hatch is OfficeChartHatchPattern hatch && shape.Width > 0 && shape.Height > 0) {
            OfficeClipPath clip;
            if (shape.Kind == OfficeShapeKind.Ellipse) {
                double rx = shape.Width / 2, ry = shape.Height / 2, k = 0.5522847498307936;
                clip = OfficeClipPath.Path(OfficePathCommand.MoveTo(rx * 2, ry),
                    OfficePathCommand.CubicBezierTo(rx * 2, ry + k * ry, rx + k * rx, ry * 2, rx, ry * 2),
                    OfficePathCommand.CubicBezierTo(rx - k * rx, ry * 2, 0, ry + k * ry, 0, ry),
                    OfficePathCommand.CubicBezierTo(0, ry - k * ry, rx - k * rx, 0, rx, 0),
                    OfficePathCommand.CubicBezierTo(rx + k * rx, 0, rx * 2, ry - k * ry, rx * 2, ry), OfficePathCommand.Close());
            } else if (shape.Kind == OfficeShapeKind.Path) clip = OfficeClipPath.Path(shape.PathCommands);
            else if (shape.Kind == OfficeShapeKind.Polygon) {
                var commands = shape.Points.Select((point, index) => index == 0 ? OfficePathCommand.MoveTo(point) : OfficePathCommand.LineTo(point)).ToList();
                commands.Add(OfficePathCommand.Close());
                clip = OfficeClipPath.Path(commands);
            } else clip = OfficeClipPath.Rectangle(shape.Width, shape.Height);
            AddPointHatch(drawing, x, y, shape.Width, shape.Height, clip, hatch, style.HatchColor!.Value);
            if (outline.HasValue) AddShape(drawing, shape, x, y, null, outline, width);
        }
    }

    private static void AddStyledPointPolygon(OfficeDrawing drawing, IReadOnlyList<OfficePoint> points,
        OfficeColor color, OfficeChartPointStyle? pointStyle, OfficeColor? defaultOutline, double defaultWidth) {
        OfficeColor? fill = pointStyle?.NoFill == true ? null : pointStyle?.FillColor ?? color;
        OfficeColor? outline = pointStyle?.ShowOutline == false ? null : pointStyle?.OutlineColor ?? defaultOutline;
        double outlineWidth = pointStyle?.OutlineWidth ?? defaultWidth;
        if (pointStyle?.ShowOutline == true && !outline.HasValue) outline = OfficeColor.Black;
        AddPolygonShape(drawing, points, fill, null, 0);
        if (pointStyle?.Hatch is OfficeChartHatchPattern hatch) {
            AddPointHatch(drawing, points, hatch, pointStyle.HatchColor!.Value);
        }
        if (outline.HasValue && outlineWidth > 0) AddPolygonShape(drawing, points, null, outline, outlineWidth);
    }

    private static void AddPointSwatch(OfficeDrawing drawing, double x, double y, double size,
        OfficeColor color, OfficeChartPointStyle? style) {
        if (style == null) {
            AddShape(drawing, OfficeShape.Rectangle(size, size), x, y, color, null, 0);
            return;
        }
        AddStyledPointPolygon(drawing, new[] { new OfficePoint(x, y), new OfficePoint(x + size, y),
            new OfficePoint(x + size, y + size), new OfficePoint(x, y + size) }, color, style, null, 0.75);
    }

    private static void AddPointHatch(OfficeDrawing drawing, IReadOnlyList<OfficePoint> points,
        OfficeChartHatchPattern hatch, OfficeColor color) {
        double left = points.Min(point => point.X), top = points.Min(point => point.Y);
        double width = points.Max(point => point.X) - left, height = points.Max(point => point.Y) - top;
        if (width <= 0 || height <= 0) return;
        var commands = new List<OfficePathCommand> { OfficePathCommand.MoveTo(points[0].X - left, points[0].Y - top) };
        for (int index = 1; index < points.Count; index++)
            commands.Add(OfficePathCommand.LineTo(points[index].X - left, points[index].Y - top));
        commands.Add(OfficePathCommand.Close());
        AddPointHatch(drawing, left, top, width, height, OfficeClipPath.Path(commands), hatch, color);
    }

    private static void AddPointHatch(OfficeDrawing drawing, double left, double top, double width, double height,
        OfficeClipPath clip, OfficeChartHatchPattern hatch, OfficeColor color) {
        var strokes = new OfficeDrawing(width, height);
        // Keep cost bounded for large chart frames while maintaining a readable normal-size hatch.
        double step = Math.Max(6, (width + height) / 512);
        void Line(double x1, double y1, double x2, double y2) =>
            AddShape(strokes, OfficeShape.Line(x1, y1, x2, y2), 0, 0, null, color, 0.75);
        if (hatch == OfficeChartHatchPattern.Horizontal || hatch == OfficeChartHatchPattern.Cross)
            for (double y = step / 2; y < height; y += step) Line(0, y, width, y);
        if (hatch == OfficeChartHatchPattern.Vertical || hatch == OfficeChartHatchPattern.Cross)
            for (double x = step / 2; x < width; x += step) Line(x, 0, x, height);
        if (hatch == OfficeChartHatchPattern.ForwardDiagonal || hatch == OfficeChartHatchPattern.DiagonalCross)
            for (double sum = step / 2; sum < width + height; sum += step) {
                double x1 = Math.Max(0, sum - height), x2 = Math.Min(width, sum);
                Line(x1, sum - x1, x2, sum - x2);
            }
        if (hatch == OfficeChartHatchPattern.BackwardDiagonal || hatch == OfficeChartHatchPattern.DiagonalCross)
            for (double offset = -height + step / 2; offset < width; offset += step) {
                double x1 = Math.Max(0, offset), x2 = Math.Min(width, offset + height);
                Line(x1, x1 - offset, x2, x2 - offset);
            }
        drawing.AddClippedDrawing(strokes, left, top, clip);
    }
}
