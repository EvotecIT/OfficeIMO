using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    private const int MaximumSvgPatternTiles = 16384;

    private static bool TryAddSvgPatternFill(
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
                out OfficeTransform patternTransform, out bool objectBoundingBoxContent)) {
            unsupported++;
            ClearShapeFill(shape.Shape);
            return false;
        }

        var tile = new OfficeDrawing(tileWidth, tileHeight);
        SvgPaintContext tileStyle = ResolveDefinitionPaintContext(pattern, paintServers, ref unsupported);
        OfficeTransform contentTransform;
        double contentViewX;
        double contentViewY;
        if (objectBoundingBoxContent) {
            contentTransform = OfficeTransform.Scale(shape.Shape.Width, shape.Shape.Height);
            contentViewX = 0D;
            contentViewY = 0D;
        } else {
            contentTransform = OfficeTransform.Identity;
            contentViewX = originX;
            contentViewY = originY;
        }
        AddChildren(pattern, tile, tileStyle, paintServers, references, contentTransform, contentViewX, contentViewY,
            maximumElements, maximumViewportDimension, maximumViewportPixels, depth + 1,
            ref visited, ref pathCommands, ref pathCommandLimitExceeded, ref unsupported);
        if (tile.Elements.Count == 0) {
            unsupported++;
            ClearShapeFill(shape.Shape);
            return false;
        }

        var repeated = new OfficeDrawing(drawing.Width, drawing.Height);
        double localOriginX = originX - viewX;
        double localOriginY = originY - viewY;
        OfficeTransform repeatTransform = patternTransform;
        try {
            repeated.AddTilingPattern(
                tile,
                new OfficeImagePlacement(0D, 0D, drawing.Width, drawing.Height),
                tileWidth,
                tileHeight,
                repeatX: true,
                repeatY: true,
                repeatTransform,
                localOriginX,
                localOriginY,
                MaximumSvgPatternTiles);
        } catch (ArgumentException) {
            unsupported++;
            ClearShapeFill(shape.Shape);
            return false;
        } catch (InvalidOperationException) {
            unsupported++;
            ClearShapeFill(shape.Shape);
            return false;
        }

        if (!TryCreateShapeClipPath(shape.Shape, out OfficeClipPath? clipPath)) {
            unsupported++;
            ClearShapeFill(shape.Shape);
            return false;
        }
        var clipped = new OfficeDrawing(drawing.Width, drawing.Height);
        try {
            clipped.AddClippedDrawing(repeated, shape.X, shape.Y, clipPath!, -shape.X, -shape.Y);
        } catch (ArgumentOutOfRangeException) {
            unsupported++;
            ClearShapeFill(shape.Shape);
            return false;
        }
        patternLayer = new OfficeDrawing(drawing.Width, drawing.Height);
        patternLayer.AddEffectDrawing(
            clipped,
            elementTransform,
            Math.Max(0D, Math.Min(1D, style.FillOpacity * style.Opacity)));
        ClearShapeFill(shape.Shape);
        return true;
    }

    private static bool TryCreateShapeClipPath(OfficeShape shape, out OfficeClipPath? clipPath) {
        clipPath = null;
        try {
            switch (shape.Kind) {
                case OfficeShapeKind.Rectangle:
                    clipPath = OfficeClipPath.Rectangle(shape.Width, shape.Height);
                    return true;
                case OfficeShapeKind.RoundedRectangle:
                    clipPath = OfficeClipPath.RoundedRectangle(shape.Width, shape.Height, shape.CornerRadius);
                    return true;
                case OfficeShapeKind.Ellipse:
                    clipPath = OfficeClipPath.Path(CreateEllipseClipCommands(shape.Width, shape.Height));
                    return true;
                case OfficeShapeKind.Polygon:
                    if (shape.Points.Count < 3) return false;
                    var polygon = new List<OfficePathCommand>(shape.Points.Count + 1);
                    for (int index = 0; index < shape.Points.Count; index++) {
                        polygon.Add(index == 0
                            ? OfficePathCommand.MoveTo(shape.Points[index])
                            : OfficePathCommand.LineTo(shape.Points[index]));
                    }
                    polygon.Add(OfficePathCommand.Close());
                    clipPath = OfficeClipPath.Path(polygon, shape.FillRule);
                    return true;
                case OfficeShapeKind.Path:
                    clipPath = OfficeClipPath.Path(shape.PathCommands, shape.FillRule);
                    return true;
                default:
                    return false;
            }
        } catch (ArgumentException) {
            return false;
        }
    }

    private static IEnumerable<OfficePathCommand> CreateEllipseClipCommands(double width, double height) {
        const int Segments = 72;
        double centerX = width / 2D;
        double centerY = height / 2D;
        for (int index = 0; index < Segments; index++) {
            double angle = Math.PI * 2D * index / Segments;
            var point = new OfficePoint(
                centerX + (Math.Cos(angle) * centerX),
                centerY + (Math.Sin(angle) * centerY));
            yield return index == 0 ? OfficePathCommand.MoveTo(point) : OfficePathCommand.LineTo(point);
        }
        yield return OfficePathCommand.Close();
    }

    private static bool TryResolveSvgPatternGeometry(
        XElement pattern,
        OfficeDrawingShape shape,
        double viewportWidth,
        double viewportHeight,
        double viewX,
        double viewY,
        out double x,
        out double y,
        out double width,
        out double height,
        out OfficeTransform transform,
        out bool objectBoundingBoxContent) {
        x = y = width = height = 0D;
        transform = OfficeTransform.Identity;
        objectBoundingBoxContent = false;
        if (pattern.Attribute("viewBox") != null || pattern.Attribute("preserveAspectRatio") != null) return false;
        string units = pattern.Attribute("patternUnits")?.Value.Trim() ?? "objectBoundingBox";
        bool userSpace = units.Equals("userSpaceOnUse", StringComparison.OrdinalIgnoreCase);
        if (!userSpace && !units.Equals("objectBoundingBox", StringComparison.OrdinalIgnoreCase)) return false;
        string contentUnits = pattern.Attribute("patternContentUnits")?.Value.Trim() ?? "userSpaceOnUse";
        objectBoundingBoxContent = contentUnits.Equals("objectBoundingBox", StringComparison.OrdinalIgnoreCase);
        if (!objectBoundingBoxContent && !contentUnits.Equals("userSpaceOnUse", StringComparison.OrdinalIgnoreCase)) return false;

        if (userSpace) {
            if (!TryPatternUserLength(pattern.Attribute("x")?.Value, viewportWidth, viewX, 0D, out x)
                || !TryPatternUserLength(pattern.Attribute("y")?.Value, viewportHeight, viewY, 0D, out y)
                || !TryPatternUserLength(pattern.Attribute("width")?.Value, viewportWidth, 0D, double.NaN, out width)
                || !TryPatternUserLength(pattern.Attribute("height")?.Value, viewportHeight, 0D, double.NaN, out height)) return false;
        } else {
            if (!TryPatternBoxFraction(pattern.Attribute("x")?.Value, 0D, out double xFraction)
                || !TryPatternBoxFraction(pattern.Attribute("y")?.Value, 0D, out double yFraction)
                || !TryPatternBoxFraction(pattern.Attribute("width")?.Value, double.NaN, out double widthFraction)
                || !TryPatternBoxFraction(pattern.Attribute("height")?.Value, double.NaN, out double heightFraction)) return false;
            x = viewX + shape.X + (xFraction * shape.Shape.Width);
            y = viewY + shape.Y + (yFraction * shape.Shape.Height);
            width = widthFraction * shape.Shape.Width;
            height = heightFraction * shape.Shape.Height;
        }
        if (width <= 0D || height <= 0D || width > viewportWidth * 4D || height > viewportHeight * 4D) return false;
        string? transformText = pattern.Attribute("patternTransform")?.Value;
        return string.IsNullOrWhiteSpace(transformText) || OfficeSvgTransformParser.TryParse(transformText, out transform);
    }

    private static bool TryPatternUserLength(string? text, double extent, double origin, double defaultValue, out double value) {
        value = defaultValue;
        if (string.IsNullOrWhiteSpace(text)) return !double.IsNaN(defaultValue);
        if (!TryViewportLength(text, extent, out value, out bool percentage)) return false;
        if (percentage) value += origin;
        return true;
    }

    private static bool TryPatternBoxFraction(string? text, double defaultValue, out double value) {
        value = defaultValue;
        if (string.IsNullOrWhiteSpace(text)) return !double.IsNaN(defaultValue);
        string normalized = text!.Trim();
        bool percentage = normalized.EndsWith("%", StringComparison.Ordinal);
        if (percentage) normalized = normalized.Substring(0, normalized.Length - 1).Trim();
        if (!double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out value)
            || double.IsNaN(value)
            || double.IsInfinity(value)) return false;
        if (percentage) value /= 100D;
        return true;
    }

    private static void ClearShapeFill(OfficeShape shape) {
        shape.FillColor = null;
        shape.FillGradient = null;
        shape.FillRadialGradient = null;
        shape.FillOpacity = null;
    }
}
