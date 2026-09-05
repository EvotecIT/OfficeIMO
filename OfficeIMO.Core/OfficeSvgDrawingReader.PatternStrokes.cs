using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    private static bool TryAddSvgPatternStroke(
        XElement? pattern,
        OfficeDrawingShape shape,
        OfficeDrawing drawing,
        SvgPaintContext style,
        SvgPaintServerRegistry paintServers,
        SvgElementReferenceRegistry references,
        OfficeTransform elementTransform,
        double viewX,
        double viewY,
        int maximumElements,
        double maximumViewportDimension,
        double maximumViewportPixels,
        int depth,
        ref int visited,
        ref int pathCommands,
        ref bool pathCommandLimitExceeded,
        ref int unsupported,
        out OfficeDrawing? patternLayer) {
        patternLayer = null;
        if (pattern == null) return false;
        if (!TryResolveSvgPatternGeometry(pattern, shape, drawing.Width, drawing.Height, viewX, viewY,
                out double originX, out double originY, out double tileWidth, out double tileHeight,
                out OfficeTransform patternTransform, out bool objectBoundingBoxContent)
            || !TryCreateShapeStrokeClipPath(shape.Shape, out OfficeClipPath? strokeClip)) {
            unsupported++;
            ClearShapeStroke(shape.Shape);
            return false;
        }

        var tile = new OfficeDrawing(tileWidth, tileHeight);
        SvgPaintContext tileStyle = ResolveDefinitionPaintContext(pattern, paintServers, ref unsupported);
        OfficeTransform contentTransform = objectBoundingBoxContent
            ? OfficeTransform.Scale(shape.Shape.Width, shape.Shape.Height)
            : OfficeTransform.Identity;
        double contentViewX = objectBoundingBoxContent ? 0D : originX;
        double contentViewY = objectBoundingBoxContent ? 0D : originY;
        AddChildren(pattern, tile, tileStyle, paintServers, references, contentTransform, contentViewX, contentViewY,
            maximumElements, maximumViewportDimension, maximumViewportPixels, depth + 1,
            ref visited, ref pathCommands, ref pathCommandLimitExceeded, ref unsupported);
        if (tile.Elements.Count == 0) {
            unsupported++;
            ClearShapeStroke(shape.Shape);
            return false;
        }

        var repeated = new OfficeDrawing(drawing.Width, drawing.Height);
        try {
            repeated.AddTilingPattern(
                tile,
                new OfficeImagePlacement(0D, 0D, drawing.Width, drawing.Height),
                tileWidth,
                tileHeight,
                repeatX: true,
                repeatY: true,
                patternTransform,
                originX - viewX,
                originY - viewY,
                MaximumSvgPatternTiles);
        } catch (ArgumentException) {
            unsupported++;
            ClearShapeStroke(shape.Shape);
            return false;
        } catch (InvalidOperationException) {
            unsupported++;
            ClearShapeStroke(shape.Shape);
            return false;
        }

        var clipped = new OfficeDrawing(drawing.Width, drawing.Height);
        try {
            clipped.AddClippedDrawing(repeated, shape.X, shape.Y, strokeClip!, -shape.X, -shape.Y);
        } catch (ArgumentOutOfRangeException) {
            unsupported++;
            ClearShapeStroke(shape.Shape);
            return false;
        }
        patternLayer = new OfficeDrawing(drawing.Width, drawing.Height);
        patternLayer.AddEffectDrawing(
            clipped,
            elementTransform,
            Math.Max(0D, Math.Min(1D, style.StrokeOpacity * style.Opacity)));
        ClearShapeStroke(shape.Shape);
        return true;
    }

    private static bool TryCreateShapeStrokeClipPath(OfficeShape shape, out OfficeClipPath? clipPath) {
        clipPath = null;
        if (shape.StrokeWidth <= 0D) return false;
        var commands = new List<OfficePathCommand>();
        foreach (OfficeFlattenedPathContour contour in GetStrokeContours(shape)) {
            AppendStrokeOutline(commands, contour.Points, contour.Closed, shape);
            if (commands.Count > MaximumSvgPathCommands) return false;
        }
        if (commands.Count == 0) return false;
        try {
            clipPath = OfficeClipPath.Path(commands, OfficeFillRule.NonZero);
            return true;
        } catch (ArgumentException) {
            return false;
        }
    }

    private static IReadOnlyList<OfficeFlattenedPathContour> GetStrokeContours(OfficeShape shape) {
        switch (shape.Kind) {
            case OfficeShapeKind.Line:
                return shape.Points.Count >= 2
                    ? new[] { new OfficeFlattenedPathContour(shape.Points, false) }
                    : Array.Empty<OfficeFlattenedPathContour>();
            case OfficeShapeKind.Rectangle:
                return new[] { new OfficeFlattenedPathContour(new[] {
                    new OfficePoint(0D, 0D), new OfficePoint(shape.Width, 0D),
                    new OfficePoint(shape.Width, shape.Height), new OfficePoint(0D, shape.Height)
                }, true) };
            case OfficeShapeKind.RoundedRectangle:
                return new[] { new OfficeFlattenedPathContour(CreateRoundedStrokeContour(shape.Width, shape.Height, shape.CornerRadius), true) };
            case OfficeShapeKind.Ellipse:
                return new[] { new OfficeFlattenedPathContour(CreateEllipseStrokeContour(shape.Width, shape.Height), true) };
            case OfficeShapeKind.Polygon:
                return new[] { new OfficeFlattenedPathContour(shape.Points, true) };
            case OfficeShapeKind.Path:
                return OfficePathFlattener.Flatten(shape.PathCommands, 0D, 0D, 1D);
            default:
                return Array.Empty<OfficeFlattenedPathContour>();
        }
    }

    private static IReadOnlyList<OfficePoint> CreateRoundedStrokeContour(double width, double height, double radius) {
        radius = Math.Max(0D, Math.Min(radius, Math.Min(width, height) / 2D));
        if (radius <= 0D) return new[] {
            new OfficePoint(0D, 0D), new OfficePoint(width, 0D),
            new OfficePoint(width, height), new OfficePoint(0D, height)
        };
        var points = new List<OfficePoint>(36);
        AppendArc(points, width - radius, radius, radius, -90D, 0D);
        AppendArc(points, width - radius, height - radius, radius, 0D, 90D);
        AppendArc(points, radius, height - radius, radius, 90D, 180D);
        AppendArc(points, radius, radius, radius, 180D, 270D);
        return points;
    }

    private static IReadOnlyList<OfficePoint> CreateEllipseStrokeContour(double width, double height) {
        var points = new List<OfficePoint>(72);
        for (int index = 0; index < 72; index++) {
            double angle = Math.PI * 2D * index / 72D;
            points.Add(new OfficePoint(
                (width / 2D) + (Math.Cos(angle) * width / 2D),
                (height / 2D) + (Math.Sin(angle) * height / 2D)));
        }
        return points;
    }

    private static void AppendArc(ICollection<OfficePoint> points, double centerX, double centerY, double radius, double startDegrees, double endDegrees) {
        for (int index = 0; index <= 8; index++) {
            double angle = OfficeGeometry.DegreesToRadians(startDegrees + ((endDegrees - startDegrees) * index / 8D));
            points.Add(new OfficePoint(centerX + (Math.Cos(angle) * radius), centerY + (Math.Sin(angle) * radius)));
        }
    }

    private static void AppendStrokeOutline(List<OfficePathCommand> commands, IReadOnlyList<OfficePoint> source, bool closed, OfficeShape shape) {
        if (source.Count < 2) return;
        var points = new List<OfficePoint>(source.Count);
        for (int index = 0; index < source.Count; index++) {
            if (index == source.Count - 1 && source[index] == source[0]) continue;
            points.Add(source[index]);
        }
        if (points.Count < 2) return;

        double half = shape.StrokeWidth / 2D;
        if (TryGetStrokeDashPattern(shape, out IReadOnlyList<double>? dashPattern)) {
            AppendDashedStrokeOutline(commands, points, closed, shape, half, dashPattern!);
            return;
        }
        int segmentCount = closed ? points.Count : points.Count - 1;
        for (int index = 0; index < segmentCount; index++) {
            OfficePoint start = points[index];
            OfficePoint end = points[(index + 1) % points.Count];
            bool extendStart = !closed && index == 0 && shape.StrokeLineCap == OfficeStrokeLineCap.Square;
            bool extendEnd = !closed && index == segmentCount - 1 && shape.StrokeLineCap == OfficeStrokeLineCap.Square;
            AppendStrokeSegment(commands, start, end, half, extendStart, extendEnd);
        }

        int firstJoin = closed ? 0 : 1;
        int lastJoin = closed ? points.Count - 1 : points.Count - 2;
        for (int index = firstJoin; index <= lastJoin; index++) {
            OfficePoint previous = points[(index - 1 + points.Count) % points.Count];
            OfficePoint vertex = points[index];
            OfficePoint next = points[(index + 1) % points.Count];
            AppendStrokeJoin(commands, previous, vertex, next, half, shape.StrokeLineJoin ?? OfficeStrokeLineJoin.Miter, shape.StrokeMiterLimit);
        }

        if (!closed && shape.StrokeLineCap == OfficeStrokeLineCap.Round) {
            AppendCircle(commands, points[0], half);
            AppendCircle(commands, points[points.Count - 1], half);
        }
    }

    private static bool TryGetStrokeDashPattern(OfficeShape shape, out IReadOnlyList<double>? pattern) {
        pattern = shape.StrokeDashArray.Count > 0
            ? shape.StrokeDashArray
            : shape.StrokeDashStyle == OfficeStrokeDashStyle.Solid
                ? null
                : shape.StrokeDashStyle.GetDashPattern(shape.StrokeWidth);
        if (pattern == null || pattern.Count == 0) return false;
        if ((pattern.Count & 1) == 0) return true;
        var doubled = new List<double>(pattern.Count * 2);
        for (int pass = 0; pass < 2; pass++) {
            for (int index = 0; index < pattern.Count; index++) doubled.Add(pattern[index]);
        }
        pattern = doubled;
        return true;
    }

    private static void AppendDashedStrokeOutline(
        List<OfficePathCommand> commands,
        IReadOnlyList<OfficePoint> points,
        bool closed,
        OfficeShape shape,
        double half,
        IReadOnlyList<double> pattern) {
        double cycle = 0D;
        for (int index = 0; index < pattern.Count; index++) cycle += pattern[index];
        if (cycle <= 0D || double.IsNaN(cycle) || double.IsInfinity(cycle)) return;
        double patternPosition = shape.StrokeDashOffset % cycle;
        if (patternPosition < 0D) patternPosition += cycle;
        int segmentCount = closed ? points.Count : points.Count - 1;
        for (int segment = 0; segment < segmentCount; segment++) {
            OfficePoint start = points[segment];
            OfficePoint end = points[(segment + 1) % points.Count];
            double dx = end.X - start.X;
            double dy = end.Y - start.Y;
            double length = Math.Sqrt((dx * dx) + (dy * dy));
            if (length <= 0.000000001D) continue;
            double consumed = 0D;
            while (consumed < length - 0.000000001D) {
                ResolveDashPosition(pattern, patternPosition, out int patternIndex, out double within);
                double available = pattern[patternIndex] - within;
                if (available <= 0.000000001D) {
                    patternPosition = AdvanceDashPosition(patternPosition, Math.Max(available, 0.000000001D), cycle);
                    continue;
                }
                double take = Math.Min(length - consumed, available);
                if ((patternIndex & 1) == 0 && take > 0.000000001D) {
                    double startRatio = consumed / length;
                    double endRatio = (consumed + take) / length;
                    var dashStart = new OfficePoint(start.X + (dx * startRatio), start.Y + (dy * startRatio));
                    var dashEnd = new OfficePoint(start.X + (dx * endRatio), start.Y + (dy * endRatio));
                    bool square = shape.StrokeLineCap == OfficeStrokeLineCap.Square;
                    AppendStrokeSegment(commands, dashStart, dashEnd, half, square, square);
                    if (shape.StrokeLineCap == OfficeStrokeLineCap.Round) {
                        AppendCircle(commands, dashStart, half);
                        AppendCircle(commands, dashEnd, half);
                    }
                }
                consumed += take;
                patternPosition = AdvanceDashPosition(patternPosition, take, cycle);
            }
        }
    }

    private static void ResolveDashPosition(IReadOnlyList<double> pattern, double position, out int index, out double within) {
        index = 0;
        within = position;
        while (index < pattern.Count - 1 && within >= pattern[index]) {
            within -= pattern[index];
            index++;
        }
    }

    private static double AdvanceDashPosition(double position, double distance, double cycle) {
        double advanced = (position + distance) % cycle;
        return advanced < 0D ? advanced + cycle : advanced;
    }

    private static void AppendStrokeSegment(List<OfficePathCommand> commands, OfficePoint start, OfficePoint end, double half, bool extendStart, bool extendEnd) {
        double dx = end.X - start.X;
        double dy = end.Y - start.Y;
        double length = Math.Sqrt((dx * dx) + (dy * dy));
        if (length <= 0.000000001D) return;
        double ux = dx / length;
        double uy = dy / length;
        double sx = start.X - (extendStart ? ux * half : 0D);
        double sy = start.Y - (extendStart ? uy * half : 0D);
        double ex = end.X + (extendEnd ? ux * half : 0D);
        double ey = end.Y + (extendEnd ? uy * half : 0D);
        double nx = -uy * half;
        double ny = ux * half;
        AppendPolygon(commands, new[] {
            new OfficePoint(sx + nx, sy + ny), new OfficePoint(ex + nx, ey + ny),
            new OfficePoint(ex - nx, ey - ny), new OfficePoint(sx - nx, sy - ny)
        });
    }

    private static void AppendStrokeJoin(List<OfficePathCommand> commands, OfficePoint previous, OfficePoint vertex, OfficePoint next, double half, OfficeStrokeLineJoin join, double miterLimit) {
        if (!TryUnitDirection(previous, vertex, out double d1x, out double d1y)
            || !TryUnitDirection(vertex, next, out double d2x, out double d2y)) return;
        double cross = (d1x * d2y) - (d1y * d2x);
        if (Math.Abs(cross) <= 0.000000001D) return;
        if (join == OfficeStrokeLineJoin.Round) {
            AppendCircle(commands, vertex, half);
            return;
        }
        double side = cross > 0D ? -1D : 1D;
        var outer1 = new OfficePoint(vertex.X + (-d1y * half * side), vertex.Y + (d1x * half * side));
        var outer2 = new OfficePoint(vertex.X + (-d2y * half * side), vertex.Y + (d2x * half * side));
        if (join == OfficeStrokeLineJoin.Miter
            && TryLineIntersection(outer1, d1x, d1y, outer2, d2x, d2y, out OfficePoint miter)
            && Distance(vertex, miter) <= half * Math.Max(1D, miterLimit)) {
            AppendPolygon(commands, new[] { outer1, miter, outer2 });
        } else {
            AppendPolygon(commands, new[] { outer1, vertex, outer2 });
        }
    }

    private static bool TryUnitDirection(OfficePoint start, OfficePoint end, out double x, out double y) {
        x = end.X - start.X;
        y = end.Y - start.Y;
        double length = Math.Sqrt((x * x) + (y * y));
        if (length <= 0.000000001D) return false;
        x /= length;
        y /= length;
        return true;
    }

    private static bool TryLineIntersection(OfficePoint first, double firstX, double firstY, OfficePoint second, double secondX, double secondY, out OfficePoint intersection) {
        double cross = (firstX * secondY) - (firstY * secondX);
        if (Math.Abs(cross) <= 0.000000001D) {
            intersection = default;
            return false;
        }
        double offsetX = second.X - first.X;
        double offsetY = second.Y - first.Y;
        double distance = ((offsetX * secondY) - (offsetY * secondX)) / cross;
        intersection = new OfficePoint(first.X + (firstX * distance), first.Y + (firstY * distance));
        return true;
    }

    private static double Distance(OfficePoint first, OfficePoint second) {
        double dx = second.X - first.X;
        double dy = second.Y - first.Y;
        return Math.Sqrt((dx * dx) + (dy * dy));
    }

    private static void AppendCircle(List<OfficePathCommand> commands, OfficePoint center, double radius) {
        var points = new OfficePoint[24];
        for (int index = 0; index < points.Length; index++) {
            double angle = Math.PI * 2D * index / points.Length;
            points[index] = new OfficePoint(center.X + (Math.Cos(angle) * radius), center.Y + (Math.Sin(angle) * radius));
        }
        AppendPolygon(commands, points);
    }

    private static void AppendPolygon(List<OfficePathCommand> commands, IReadOnlyList<OfficePoint> points) {
        if (points.Count < 3) return;
        double area = 0D;
        for (int index = 0; index < points.Count; index++) {
            OfficePoint current = points[index];
            OfficePoint next = points[(index + 1) % points.Count];
            area += (current.X * next.Y) - (next.X * current.Y);
        }
        bool reverse = area < 0D;
        for (int offset = 0; offset < points.Count; offset++) {
            int index = reverse ? points.Count - 1 - offset : offset;
            commands.Add(offset == 0 ? OfficePathCommand.MoveTo(points[index]) : OfficePathCommand.LineTo(points[index]));
        }
        commands.Add(OfficePathCommand.Close());
    }

    private static void ClearShapeStroke(OfficeShape shape) {
        shape.StrokeColor = null;
        shape.StrokeGradient = null;
        shape.StrokeRadialGradient = null;
        shape.StrokeOpacity = null;
    }
}
