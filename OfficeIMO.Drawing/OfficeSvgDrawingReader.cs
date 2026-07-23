using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Drawing;

/// <summary>
/// Reads a bounded subset of SVG into the shared dependency-free drawing scene.
/// </summary>
public static partial class OfficeSvgDrawingReader {
    private const int MaximumInputBytes = 8 * 1024 * 1024;
    private const int MaximumSvgNestingDepth = 128;
    private const int MaximumSvgPathCommands = 20000;
    private const double MaximumSvgTransformCoefficient = 1024D;
    private const double MaximumSvgTransformOffset = 1000000D;

    /// <summary>Attempts to interpret supported SVG vector primitives as a shared drawing.</summary>
    public static bool TryRead(byte[]? bytes, out OfficeDrawing? drawing) =>
        TryRead(bytes, options: null, out drawing, out _);

    /// <summary>
    /// Attempts to interpret supported SVG vector primitives as a shared drawing and reports the
    /// number of elements or declarations that required omission or fallback.
    /// </summary>
    public static bool TryRead(byte[]? bytes, out OfficeDrawing? drawing, out int unsupportedFeatureCount) =>
        TryRead(bytes, options: null, out drawing, out unsupportedFeatureCount);

    /// <summary>Attempts to interpret supported SVG vector primitives using explicit bounded import options.</summary>
    public static bool TryRead(byte[]? bytes, OfficeSvgDrawingReaderOptions? options, out OfficeDrawing? drawing) =>
        TryRead(bytes, options, out drawing, out _);

    /// <summary>
    /// Attempts to interpret supported SVG vector primitives using explicit bounded import options and reports the
    /// number of elements or declarations that required omission or fallback.
    /// </summary>
    public static bool TryRead(
        byte[]? bytes,
        OfficeSvgDrawingReaderOptions? options,
        out OfficeDrawing? drawing,
        out int unsupportedFeatureCount) {
        drawing = null;
        unsupportedFeatureCount = 0;
        if (bytes == null || bytes.Length == 0 || bytes.Length > MaximumInputBytes) return false;
        int maximumElements = options?.MaximumElements ?? OfficeSvgDrawingReaderOptions.DefaultMaximumElements;
        double maximumViewportDimension = options?.MaximumViewportDimension ?? OfficeSvgDrawingReaderOptions.DefaultMaximumViewportDimension;
        double maximumViewportPixels = options?.MaximumViewportPixels ?? OfficeSvgDrawingReaderOptions.DefaultMaximumViewportPixels;
        if (maximumElements <= 0 || maximumElements > OfficeSvgDrawingReaderOptions.MaximumAllowedElements) return false;
        if (maximumViewportDimension <= 0D || maximumViewportDimension > OfficeSvgDrawingReaderOptions.MaximumAllowedViewportDimension ||
            maximumViewportPixels <= 0D || maximumViewportPixels > OfficeSvgDrawingReaderOptions.MaximumAllowedViewportPixels) return false;

        try {
            var settings = new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                MaxCharactersInDocument = MaximumInputBytes
            };
            XDocument document;
            using (var stream = new MemoryStream(bytes, writable: false))
            using (XmlReader reader = XmlReader.Create(stream, settings)) {
                document = XDocument.Load(reader, LoadOptions.None);
            }

            XElement? root = document.Root;
            if (root == null || !string.Equals(root.Name.LocalName, "svg", StringComparison.OrdinalIgnoreCase)) return false;
            if (root.Descendants().Take(maximumElements + 1).Count() > maximumElements) return false;
            if (!TryResolveViewport(bytes, root, maximumViewportDimension, maximumViewportPixels,
                    out double viewX, out double viewY, out double width, out double height)) return false;

            var result = new OfficeDrawing(width, height);
            int visited = 0;
            int pathCommands = 0;
            SvgDefinitionRegistry definitions = SvgDefinitionRegistry.Create(root);
            var paintServers = new SvgPaintServerRegistry(definitions);
            var references = new SvgElementReferenceRegistry(definitions);
            var context = ResolvePaintContext(root, SvgPaintContext.Default, paintServers, ref unsupportedFeatureCount);
            OfficeTransform rootTransform = ResolveTransform(root, OfficeTransform.Identity, viewX, viewY, ref unsupportedFeatureCount);
            AddChildren(root, result, context, paintServers, references, rootTransform, viewX, viewY,
                maximumElements, maximumViewportDimension, maximumViewportPixels, 0,
                ref visited, ref pathCommands, ref unsupportedFeatureCount);
            if (visited > maximumElements) return false;
            drawing = result;
            return IsSupportedSvgViewport(width, height, maximumViewportDimension, maximumViewportPixels);
        } catch (XmlException) {
            return false;
        } catch (InvalidOperationException) {
            return false;
        } catch (ArgumentException) {
            return false;
        }
    }

    private static bool TryResolveViewport(
        byte[] bytes,
        XElement root,
        double maximumViewportDimension,
        double maximumViewportPixels,
        out double viewX,
        out double viewY,
        out double width,
        out double height) {
        viewX = viewY = 0D;
        width = height = 0D;
        if (TryParseNumberList(root.Attribute("viewBox")?.Value, out IReadOnlyList<double> viewBox)
            && viewBox.Count == 4
            && viewBox[2] > 0D
            && viewBox[3] > 0D) {
            viewX = viewBox[0];
            viewY = viewBox[1];
            width = viewBox[2];
            height = viewBox[3];
            return IsSupportedSvgViewport(width, height, maximumViewportDimension, maximumViewportPixels);
        }

        if (!OfficeImageReader.TryIdentify(bytes, ".svg", out OfficeImageInfo info) || info.Width <= 0 || info.Height <= 0) return false;
        width = info.Width * 96D / Math.Max(1D, info.DpiX);
        height = info.Height * 96D / Math.Max(1D, info.DpiY);
        return IsSupportedSvgViewport(width, height, maximumViewportDimension, maximumViewportPixels);
    }

    private static bool IsSupportedSvgViewport(
        double width,
        double height,
        double maximumViewportDimension,
        double maximumViewportPixels) =>
        width > 0D && height > 0D &&
        width <= maximumViewportDimension && height <= maximumViewportDimension &&
        width * height <= maximumViewportPixels;

    private static void AddChildren(
        XElement parent,
        OfficeDrawing drawing,
        SvgPaintContext inherited,
        SvgPaintServerRegistry paintServers,
        SvgElementReferenceRegistry references,
        OfficeTransform inheritedTransform,
        double viewX,
        double viewY,
        int maximumElements,
        double maximumViewportDimension,
        double maximumViewportPixels,
        int depth,
        ref int visited,
        ref int pathCommands,
        ref int unsupported) {
        if (depth > MaximumSvgNestingDepth) {
            unsupported++;
            return;
        }
        foreach (XElement element in parent.Elements()) {
            AddElement(element, drawing, inherited, paintServers, references, inheritedTransform, viewX, viewY,
                maximumElements, maximumViewportDimension, maximumViewportPixels, depth,
                ref visited, ref pathCommands, ref unsupported);
            if (visited > maximumElements) return;
        }
    }

    private static void AddElement(
        XElement element,
        OfficeDrawing drawing,
        SvgPaintContext inherited,
        SvgPaintServerRegistry paintServers,
        SvgElementReferenceRegistry references,
        OfficeTransform inheritedTransform,
        double viewX,
        double viewY,
        int maximumElements,
        double maximumViewportDimension,
        double maximumViewportPixels,
        int depth,
        ref int visited,
        ref int pathCommands,
        ref int unsupported) {
        visited++;
        if (visited > maximumElements) return;
        string name = element.Name.LocalName.ToLowerInvariant();
        if (name is "title" or "desc" or "metadata" or "lineargradient" or "radialgradient" or "stop") return;
        if (name == "defs") return;

        SvgPaintContext style = ResolvePaintContext(element, inherited, paintServers, ref unsupported);
        if (!style.Visible) return;
        OfficeTransform transform = ResolveTransform(element, inheritedTransform, viewX, viewY, ref unsupported);
        if (name is "g" or "svg" or "a" or "switch") {
            AddChildren(element, drawing, style, paintServers, references, transform, viewX, viewY,
                maximumElements, maximumViewportDimension, maximumViewportPixels, depth + 1,
                ref visited, ref pathCommands, ref unsupported);
            return;
        }
        if (name == "use") {
            AddReferencedElement(element, drawing, style, paintServers, references, transform, viewX, viewY,
                maximumElements, maximumViewportDimension, maximumViewportPixels, depth + 1,
                ref visited, ref pathCommands, ref unsupported);
            return;
        }
        if (name == "text") {
            AddText(element, drawing, style, paintServers, transform, viewX, viewY, ref unsupported);
            return;
        }

        OfficeDrawingShape? shape = name switch {
            "rect" => CreateRectangle(element, style, viewX, viewY, drawing.Width, drawing.Height, ref unsupported),
            "circle" => CreateCircle(element, style, viewX, viewY, drawing.Width, drawing.Height),
            "ellipse" => CreateEllipse(element, style, viewX, viewY, drawing.Width, drawing.Height),
            "line" => CreateLine(element, style, viewX, viewY, drawing.Width, drawing.Height),
            "polygon" => CreatePolygon(element, style, viewX, viewY, close: true, ref pathCommands),
            "polyline" => CreatePolygon(element, style, viewX, viewY, close: false, ref pathCommands),
            "path" => CreatePath(element, style, viewX, viewY, ref pathCommands),
            _ => null
        };
        if (shape == null) {
            unsupported++;
            return;
        }

        ApplyDeferredPaint(shape.Shape, style, shape.X, shape.Y, drawing.Width, drawing.Height, viewX, viewY, ref unsupported);

        ApplyTransform(shape, transform);

        try {
            drawing.AddShape(shape.Shape, shape.X, shape.Y);
        } catch (ArgumentOutOfRangeException) {
            unsupported++;
        }
    }

    private static OfficeTransform ResolveTransform(
        XElement element,
        OfficeTransform inherited,
        double viewX,
        double viewY,
        ref int unsupported) {
        string? value = element.Attribute("transform")?.Value;
        if (string.IsNullOrWhiteSpace(value)) return inherited;
        if (!OfficeSvgTransformParser.TryParse(value, out OfficeTransform parsed)) {
            unsupported++;
            return inherited;
        }
        OfficeTransform normalized = OfficeTransform.Translate(viewX, viewY)
            .Then(parsed)
            .Then(OfficeTransform.Translate(-viewX, -viewY));
        OfficeTransform combined = normalized.Then(inherited);
        if (!IsSupportedSvgTransform(combined)) {
            unsupported++;
            return inherited;
        }
        return combined;
    }

    private static bool IsSupportedSvgTransform(OfficeTransform transform) =>
        Math.Abs(transform.M11) <= MaximumSvgTransformCoefficient &&
        Math.Abs(transform.M12) <= MaximumSvgTransformCoefficient &&
        Math.Abs(transform.M21) <= MaximumSvgTransformCoefficient &&
        Math.Abs(transform.M22) <= MaximumSvgTransformCoefficient &&
        Math.Abs(transform.OffsetX) <= MaximumSvgTransformOffset &&
        Math.Abs(transform.OffsetY) <= MaximumSvgTransformOffset;

    private static void ApplyTransform(OfficeDrawingShape drawingShape, OfficeTransform transform) {
        if (transform == OfficeTransform.Identity) return;
        OfficeTransform local = OfficeTransform.Translate(drawingShape.X, drawingShape.Y)
            .Then(transform)
            .Then(OfficeTransform.Translate(-drawingShape.X, -drawingShape.Y));
        OfficeShape shape = drawingShape.Shape;
        shape.Transform = shape.Transform.HasValue ? shape.Transform.Value.Then(local) : local;
    }

    private static OfficeDrawingShape? CreateRectangle(
        XElement element,
        SvgPaintContext style,
        double viewX,
        double viewY,
        double viewportWidth,
        double viewportHeight,
        ref int unsupported) {
        if (!TryViewportLength(element, "width", viewportWidth, out double width)
            || !TryViewportLength(element, "height", viewportHeight, out double height)
            || width <= 0D
            || height <= 0D) return null;
        double x = ReadViewportCoordinate(element, "x", viewX, viewportWidth);
        double y = ReadViewportCoordinate(element, "y", viewY, viewportHeight);
        double rx = ReadViewportLength(element, "rx", viewportWidth);
        double ry = ReadViewportLength(element, "ry", viewportHeight);
        if (rx <= 0D && ry > 0D) rx = ry;
        if (ry <= 0D && rx > 0D) ry = rx;
        OfficeShape shape;
        if (rx > 0D || ry > 0D) {
            if (Math.Abs(rx - ry) > 0.0001D) unsupported++;
            shape = OfficeShape.RoundedRectangle(width, height, Math.Min(Math.Min(rx, ry), Math.Min(width, height) / 2D));
        } else {
            shape = OfficeShape.Rectangle(width, height);
        }
        ApplyPaint(shape, style);
        return new OfficeDrawingShape(shape, x, y);
    }

    private static OfficeDrawingShape? CreateCircle(
        XElement element,
        SvgPaintContext style,
        double viewX,
        double viewY,
        double viewportWidth,
        double viewportHeight) {
        double normalizedDiagonal = Math.Sqrt((viewportWidth * viewportWidth) + (viewportHeight * viewportHeight)) / Math.Sqrt(2D);
        if (!TryViewportLength(element, "r", normalizedDiagonal, out double radius) || radius <= 0D) return null;
        double x = ReadViewportCoordinate(element, "cx", viewX, viewportWidth) - radius;
        double y = ReadViewportCoordinate(element, "cy", viewY, viewportHeight) - radius;
        OfficeShape shape = OfficeShape.Ellipse(radius * 2D, radius * 2D);
        ApplyPaint(shape, style);
        return new OfficeDrawingShape(shape, x, y);
    }

    private static OfficeDrawingShape? CreateEllipse(
        XElement element,
        SvgPaintContext style,
        double viewX,
        double viewY,
        double viewportWidth,
        double viewportHeight) {
        if (!TryViewportLength(element, "rx", viewportWidth, out double radiusX)
            || !TryViewportLength(element, "ry", viewportHeight, out double radiusY)
            || radiusX <= 0D
            || radiusY <= 0D) return null;
        double x = ReadViewportCoordinate(element, "cx", viewX, viewportWidth) - radiusX;
        double y = ReadViewportCoordinate(element, "cy", viewY, viewportHeight) - radiusY;
        OfficeShape shape = OfficeShape.Ellipse(radiusX * 2D, radiusY * 2D);
        ApplyPaint(shape, style);
        return new OfficeDrawingShape(shape, x, y);
    }

    private static OfficeDrawingShape? CreateLine(
        XElement element,
        SvgPaintContext style,
        double viewX,
        double viewY,
        double viewportWidth,
        double viewportHeight) {
        double x1 = ReadViewportCoordinate(element, "x1", viewX, viewportWidth);
        double y1 = ReadViewportCoordinate(element, "y1", viewY, viewportHeight);
        double x2 = ReadViewportCoordinate(element, "x2", viewX, viewportWidth);
        double y2 = ReadViewportCoordinate(element, "y2", viewY, viewportHeight);
        if (Math.Abs(x1 - x2) <= 0.0001D && Math.Abs(y1 - y2) <= 0.0001D) return null;
        OfficeShape shape = OfficeShape.Line(x1, y1, x2, y2);
        shape.FillColor = null;
        shape.StrokeColor = style.Stroke ?? style.Fill ?? OfficeColor.Black;
        shape.StrokeGradient = style.StrokeGradient;
        shape.StrokeRadialGradient = style.StrokeRadialGradient;
        shape.StrokeWidth = style.StrokeWidth;
        shape.StrokeOpacity = style.StrokeOpacity * style.Opacity;
        shape.StrokeDashStyle = style.DashStyle;
        shape.StrokeLineCap = style.LineCap;
        shape.StrokeLineJoin = style.LineJoin;
        double x = Math.Min(x1, x2);
        double y = Math.Min(y1, y2);
        return new OfficeDrawingShape(shape, x, y);
    }

    private static OfficeDrawingShape? CreatePolygon(XElement element, SvgPaintContext style, double viewX,
        double viewY, bool close, ref int pathCommands) {
        int remainingCommands = MaximumSvgPathCommands - pathCommands;
        if (remainingCommands <= 0) {
            pathCommands = MaximumSvgPathCommands;
            return null;
        }
        bool parsed = TryParseNumberList(element.Attribute("points")?.Value, remainingCommands * 2,
            out IReadOnlyList<double> values, out bool limitExceeded);
        if (!parsed || values.Count < 4 || values.Count % 2 != 0) {
            if (limitExceeded) {
                pathCommands = MaximumSvgPathCommands;
            } else if (values.Count > 0) {
                int parsedCommands = Math.Max(1, (values.Count + 1) / 2);
                pathCommands += Math.Min(remainingCommands, parsedCommands);
            }
            return null;
        }
        int commandCount = values.Count / 2;
        if (close) commandCount++;
        if (close && values.Count < 6) {
            return null;
        }
        if (commandCount > remainingCommands) {
            pathCommands = MaximumSvgPathCommands;
            return null;
        }
        var points = new List<OfficePoint>(values.Count / 2);
        for (int index = 0; index < values.Count; index += 2) points.Add(new OfficePoint(values[index] - viewX, values[index + 1] - viewY));
        double minX = points.Min(point => point.X);
        double minY = points.Min(point => point.Y);
        OfficeShape shape;
        if (close) {
            shape = OfficeShape.Polygon(points);
        } else {
            var commands = new List<OfficePathCommand> { OfficePathCommand.MoveTo(points[0]) };
            for (int index = 1; index < points.Count; index++) commands.Add(OfficePathCommand.LineTo(points[index]));
            try {
                shape = OfficeShape.Path(commands);
            } catch (ArgumentException) {
                return null;
            }
            shape.FillColor = null;
        }
        pathCommands += commandCount;
        ApplyPaint(shape, style);
        return new OfficeDrawingShape(shape, minX, minY);
    }

    private static OfficeDrawingShape? CreatePath(XElement element, SvgPaintContext style, double viewX,
        double viewY, ref int pathCommands) {
        int remaining = MaximumSvgPathCommands - pathCommands;
        if (remaining <= 0) return null;
        if (!OfficeSvgPathDataParser.TryParse(element.Attribute("d")?.Value, remaining,
                out IReadOnlyList<OfficePathCommand> parsed, out bool commandLimitExceeded)) {
            if (commandLimitExceeded) pathCommands = MaximumSvgPathCommands;
            return null;
        }
        pathCommands += parsed.Count;
        var commands = new List<OfficePathCommand>(parsed.Count + 1);
        double minX = double.PositiveInfinity;
        double minY = double.PositiveInfinity;
        double maxX = double.NegativeInfinity;
        double maxY = double.NegativeInfinity;
        foreach (OfficePathCommand source in parsed) {
            OfficePathCommand command = source.Translate(viewX, viewY);
            commands.Add(command);
            IncludeCommandBounds(command, ref minX, ref minY, ref maxX, ref maxY);
        }
        if (double.IsInfinity(minX) || double.IsInfinity(minY)) return null;
        if (maxX - minX <= 0.0001D) commands.Add(OfficePathCommand.MoveTo(maxX + 0.0001D, maxY));
        if (maxY - minY <= 0.0001D) commands.Add(OfficePathCommand.MoveTo(maxX, maxY + 0.0001D));
        OfficeShape shape;
        try {
            shape = OfficeShape.Path(commands);
        } catch (ArgumentException) {
            return null;
        }
        ApplyPaint(shape, style);
        return new OfficeDrawingShape(shape, minX, minY);
    }

    private static void IncludeCommandBounds(
        OfficePathCommand command,
        ref double minX,
        ref double minY,
        ref double maxX,
        ref double maxY) {
        switch (command.Kind) {
            case OfficePathCommandKind.MoveTo:
            case OfficePathCommandKind.LineTo:
                IncludePoint(command.Point, ref minX, ref minY, ref maxX, ref maxY);
                break;
            case OfficePathCommandKind.QuadraticBezierTo:
                IncludePoint(command.ControlPoint1, ref minX, ref minY, ref maxX, ref maxY);
                IncludePoint(command.Point, ref minX, ref minY, ref maxX, ref maxY);
                break;
            case OfficePathCommandKind.CubicBezierTo:
                IncludePoint(command.ControlPoint1, ref minX, ref minY, ref maxX, ref maxY);
                IncludePoint(command.ControlPoint2, ref minX, ref minY, ref maxX, ref maxY);
                IncludePoint(command.Point, ref minX, ref minY, ref maxX, ref maxY);
                break;
        }
    }

    private static void IncludePoint(
        OfficePoint point,
        ref double minX,
        ref double minY,
        ref double maxX,
        ref double maxY) {
        minX = Math.Min(minX, point.X);
        minY = Math.Min(minY, point.Y);
        maxX = Math.Max(maxX, point.X);
        maxY = Math.Max(maxY, point.Y);
    }

    private static SvgPaintContext ResolvePaintContext(XElement element, SvgPaintContext inherited, SvgPaintServerRegistry paintServers, ref int unsupported) {
        SvgPaintContext result = inherited;
        ApplyProperty("color", element.Attribute("color")?.Value, paintServers, ref result, ref unsupported);
        string? styleText = element.Attribute("style")?.Value;
        string[] declarations = string.IsNullOrWhiteSpace(styleText) ? Array.Empty<string>() : styleText!.Split(';');
        foreach (string declaration in declarations) {
            int colon = declaration.IndexOf(':');
            if (colon <= 0 || !declaration.Substring(0, colon).Trim().Equals("color", StringComparison.OrdinalIgnoreCase)) continue;
            ApplyProperty("color", declaration.Substring(colon + 1).Trim(), paintServers, ref result, ref unsupported);
        }
        ApplyProperty("fill", element.Attribute("fill")?.Value, paintServers, ref result, ref unsupported);
        ApplyProperty("stroke", element.Attribute("stroke")?.Value, paintServers, ref result, ref unsupported);
        ApplyProperty("stroke-width", element.Attribute("stroke-width")?.Value, paintServers, ref result, ref unsupported);
        ApplyProperty("opacity", element.Attribute("opacity")?.Value, paintServers, ref result, ref unsupported);
        ApplyProperty("fill-opacity", element.Attribute("fill-opacity")?.Value, paintServers, ref result, ref unsupported);
        ApplyProperty("stroke-opacity", element.Attribute("stroke-opacity")?.Value, paintServers, ref result, ref unsupported);
        ApplyProperty("stroke-dasharray", element.Attribute("stroke-dasharray")?.Value, paintServers, ref result, ref unsupported);
        ApplyProperty("stroke-linecap", element.Attribute("stroke-linecap")?.Value, paintServers, ref result, ref unsupported);
        ApplyProperty("stroke-linejoin", element.Attribute("stroke-linejoin")?.Value, paintServers, ref result, ref unsupported);
        ApplyProperty("fill-rule", element.Attribute("fill-rule")?.Value, paintServers, ref result, ref unsupported);
        ApplyProperty("font-family", element.Attribute("font-family")?.Value, paintServers, ref result, ref unsupported);
        ApplyProperty("font-size", element.Attribute("font-size")?.Value, paintServers, ref result, ref unsupported);
        ApplyProperty("font-style", element.Attribute("font-style")?.Value, paintServers, ref result, ref unsupported);
        ApplyProperty("font-weight", element.Attribute("font-weight")?.Value, paintServers, ref result, ref unsupported);
        ApplyProperty("text-anchor", element.Attribute("text-anchor")?.Value, paintServers, ref result, ref unsupported);
        ApplyProperty("dominant-baseline", element.Attribute("dominant-baseline")?.Value, paintServers, ref result, ref unsupported);
        ApplyProperty("display", element.Attribute("display")?.Value, paintServers, ref result, ref unsupported);
        ApplyProperty("visibility", element.Attribute("visibility")?.Value, paintServers, ref result, ref unsupported);
        foreach (string declaration in declarations) {
            int colon = declaration.IndexOf(':');
            if (colon <= 0) continue;
            string name = declaration.Substring(0, colon).Trim();
            if (name.Equals("color", StringComparison.OrdinalIgnoreCase)) continue;
            ApplyProperty(name, declaration.Substring(colon + 1).Trim(), paintServers, ref result, ref unsupported);
        }
        return result;
    }

    private static void ApplyProperty(string name, string? value, SvgPaintServerRegistry paintServers, ref SvgPaintContext style, ref int unsupported) {
        if (string.IsNullOrWhiteSpace(value)) return;
        string normalized = value!.Trim();
        switch (name.Trim().ToLowerInvariant()) {
            case "color":
                if (normalized.Equals("currentcolor", StringComparison.OrdinalIgnoreCase)) break;
                if (!TrySvgColor(normalized, out OfficeColor currentColor)) unsupported++;
                else style.Color = currentColor;
                break;
            case "fill":
                if (!TryPaint(normalized, paintServers, style.Color, out SvgResolvedPaint fill)) {
                    unsupported++;
                    if (normalized.StartsWith("url(", StringComparison.OrdinalIgnoreCase)) style.SetFill(default);
                }
                else style.SetFill(fill);
                break;
            case "stroke":
                if (!TryPaint(normalized, paintServers, style.Color, out SvgResolvedPaint stroke)) {
                    unsupported++;
                    if (normalized.StartsWith("url(", StringComparison.OrdinalIgnoreCase)) style.SetStroke(default);
                }
                else style.SetStroke(stroke);
                break;
            case "stroke-width":
                if (!TrySvgLength(normalized, out double strokeWidth) || strokeWidth < 0D) unsupported++;
                else style.StrokeWidth = strokeWidth;
                break;
            case "opacity":
                if (!TryUnit(normalized, out double opacity)) unsupported++;
                else style.Opacity *= opacity;
                break;
            case "fill-opacity":
                if (!TryUnit(normalized, out double fillOpacity)) unsupported++;
                else style.FillOpacity = fillOpacity;
                break;
            case "stroke-opacity":
                if (!TryUnit(normalized, out double strokeOpacity)) unsupported++;
                else style.StrokeOpacity = strokeOpacity;
                break;
            case "stroke-dasharray":
                if (normalized.Equals("none", StringComparison.OrdinalIgnoreCase)) style.DashStyle = OfficeStrokeDashStyle.Solid;
                else if (TryParseNumberList(normalized, out IReadOnlyList<double> dash) && dash.Count >= 2) style.DashStyle = OfficeStrokeDashStyle.Dash;
                else unsupported++;
                break;
            case "stroke-linecap":
                if (normalized.Equals("butt", StringComparison.OrdinalIgnoreCase)) style.LineCap = OfficeStrokeLineCap.Butt;
                else if (normalized.Equals("round", StringComparison.OrdinalIgnoreCase)) style.LineCap = OfficeStrokeLineCap.Round;
                else if (normalized.Equals("square", StringComparison.OrdinalIgnoreCase)) style.LineCap = OfficeStrokeLineCap.Square;
                else unsupported++;
                break;
            case "stroke-linejoin":
                if (normalized.Equals("miter", StringComparison.OrdinalIgnoreCase)) style.LineJoin = OfficeStrokeLineJoin.Miter;
                else if (normalized.Equals("round", StringComparison.OrdinalIgnoreCase)) style.LineJoin = OfficeStrokeLineJoin.Round;
                else if (normalized.Equals("bevel", StringComparison.OrdinalIgnoreCase)) style.LineJoin = OfficeStrokeLineJoin.Bevel;
                else unsupported++;
                break;
            case "fill-rule":
                if (normalized.Equals("nonzero", StringComparison.OrdinalIgnoreCase)) style.FillRule = OfficeFillRule.NonZero;
                else if (normalized.Equals("evenodd", StringComparison.OrdinalIgnoreCase)) style.FillRule = OfficeFillRule.EvenOdd;
                else unsupported++;
                break;
            case "font-family":
                string family = normalized.Split(',')[0].Trim().Trim('\'', '"');
                if (family.Length == 0) unsupported++;
                else style.FontFamily = family;
                break;
            case "font-size":
                if (!TrySvgLength(normalized, out double fontSize) || fontSize <= 0D) unsupported++;
                else style.FontSize = fontSize;
                break;
            case "font-style":
                if (normalized.Equals("normal", StringComparison.OrdinalIgnoreCase)) style.FontStyle &= ~OfficeFontStyle.Italic;
                else if (normalized.Equals("italic", StringComparison.OrdinalIgnoreCase) || normalized.Equals("oblique", StringComparison.OrdinalIgnoreCase)) style.FontStyle |= OfficeFontStyle.Italic;
                else unsupported++;
                break;
            case "font-weight":
                if (normalized.Equals("normal", StringComparison.OrdinalIgnoreCase) || normalized == "400") style.FontStyle &= ~OfficeFontStyle.Bold;
                else if (normalized.Equals("bold", StringComparison.OrdinalIgnoreCase) || normalized.Equals("bolder", StringComparison.OrdinalIgnoreCase)) style.FontStyle |= OfficeFontStyle.Bold;
                else if (int.TryParse(normalized, NumberStyles.Integer, CultureInfo.InvariantCulture, out int weight) && weight >= 1 && weight <= 1000) {
                    if (weight >= 600) style.FontStyle |= OfficeFontStyle.Bold;
                    else style.FontStyle &= ~OfficeFontStyle.Bold;
                }
                else unsupported++;
                break;
            case "text-anchor":
                string anchor = normalized.ToLowerInvariant();
                if (anchor is "start" or "middle" or "end") style.TextAnchor = anchor;
                else unsupported++;
                break;
            case "dominant-baseline":
                string baseline = normalized.ToLowerInvariant();
                if (baseline is "auto" or "alphabetic") style.DominantBaseline = SvgDominantBaseline.Alphabetic;
                else if (baseline is "hanging" or "text-before-edge") style.DominantBaseline = SvgDominantBaseline.Hanging;
                else if (baseline is "middle" or "central") style.DominantBaseline = SvgDominantBaseline.Middle;
                else if (baseline is "text-after-edge" or "ideographic") style.DominantBaseline = SvgDominantBaseline.TextAfterEdge;
                else unsupported++;
                break;
            case "display":
                if (normalized.Equals("none", StringComparison.OrdinalIgnoreCase)) style.Visible = false;
                break;
            case "visibility":
                if (normalized.Equals("hidden", StringComparison.OrdinalIgnoreCase) || normalized.Equals("collapse", StringComparison.OrdinalIgnoreCase)) style.Visible = false;
                break;
            case "transform":
            case "filter":
            case "mask":
            case "clip-path":
            case "marker-start":
            case "marker-mid":
            case "marker-end":
                unsupported++;
                break;
        }
    }

    private static void ApplyPaint(OfficeShape shape, SvgPaintContext style) {
        shape.FillColor = style.Fill;
        shape.FillGradient = style.FillGradient;
        shape.FillRadialGradient = style.FillRadialGradient;
        shape.StrokeColor = style.Stroke;
        shape.StrokeGradient = style.StrokeGradient;
        shape.StrokeRadialGradient = style.StrokeRadialGradient;
        shape.StrokeWidth = style.StrokeWidth;
        shape.FillOpacity = style.FillOpacity * style.Opacity;
        shape.StrokeOpacity = style.StrokeOpacity * style.Opacity;
        shape.StrokeDashStyle = style.DashStyle;
        shape.StrokeLineCap = style.LineCap;
        shape.StrokeLineJoin = style.LineJoin;
        shape.FillRule = style.FillRule;
    }

    private static void ApplyDeferredPaint(
        OfficeShape shape,
        SvgPaintContext style,
        double shapeX,
        double shapeY,
        double viewportWidth,
        double viewportHeight,
        double viewX,
        double viewY,
        ref int unsupported) {
        if (style.FillDeferredGradient != null) {
            if (style.FillDeferredGradient.TryCreateForShape(shape, shapeX, shapeY, viewportWidth, viewportHeight, viewX, viewY, out OfficeLinearGradient? linear, out OfficeRadialGradient? radial)) {
                shape.FillGradient = linear;
                shape.FillRadialGradient = radial;
            } else {
                unsupported++;
            }
        }
        if (style.StrokeDeferredGradient != null) {
            if (style.StrokeDeferredGradient.TryCreateForShape(shape, shapeX, shapeY, viewportWidth, viewportHeight, viewX, viewY, out OfficeLinearGradient? linear, out OfficeRadialGradient? radial)) {
                shape.StrokeGradient = linear;
                shape.StrokeRadialGradient = radial;
            } else {
                unsupported++;
            }
        }
    }

    private static double ReadLength(XElement element, string name) => TryLength(element, name, out double value) ? value : 0D;
    private static bool TryLength(XElement element, string name, out double value) => TrySvgLength(element.Attribute(name)?.Value, out value);

    private static double ReadViewportCoordinate(XElement element, string name, double origin, double extent) {
        string? text = element.Attribute(name)?.Value;
        if (!TryViewportLength(text, extent, out double value, out _)) return -origin;
        return value - origin;
    }

    private static double ReadViewportLength(XElement element, string name, double extent) =>
        TryViewportLength(element, name, extent, out double value) ? value : 0D;

    private static bool TryViewportLength(XElement element, string name, double extent, out double value) =>
        TryViewportLength(element.Attribute(name)?.Value, extent, out value, out _);

    private static bool TryViewportLength(string? text, double extent, out double value, out bool percentage) {
        value = 0D;
        percentage = false;
        if (string.IsNullOrWhiteSpace(text)) return false;
        string normalized = text!.Trim();
        percentage = normalized.EndsWith("%", StringComparison.Ordinal);
        if (percentage) normalized = normalized.Substring(0, normalized.Length - 1).Trim();
        else if (normalized.EndsWith("px", StringComparison.OrdinalIgnoreCase)) normalized = normalized.Substring(0, normalized.Length - 2).Trim();
        if (!double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
            || double.IsNaN(parsed)
            || double.IsInfinity(parsed)) return false;
        value = percentage ? parsed * extent / 100D : parsed;
        return !double.IsNaN(value) && !double.IsInfinity(value);
    }

    private static double ReadFirstLength(XElement element, string name) {
        string? value = element.Attribute(name)?.Value;
        if (string.IsNullOrWhiteSpace(value)) return 0D;
        int separator = value!.IndexOfAny(new[] { ' ', '\t', '\r', '\n', ',' });
        return TrySvgLength(separator < 0 ? value : value.Substring(0, separator), out double parsed) ? parsed : 0D;
    }

    private static bool TrySvgLength(string? value, out double result) {
        result = 0D;
        if (string.IsNullOrWhiteSpace(value)) return false;
        string text = value!.Trim();
        if (text.EndsWith("px", StringComparison.OrdinalIgnoreCase)) text = text.Substring(0, text.Length - 2).Trim();
        return double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out result)
            && !double.IsNaN(result)
            && !double.IsInfinity(result);
    }

    private static bool TryUnit(string value, out double result) =>
        double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out result)
        && !double.IsNaN(result)
        && !double.IsInfinity(result)
        && result >= 0D
        && result <= 1D;

    private static bool TryParseNumberList(string? value, out IReadOnlyList<double> values) =>
        TryParseNumberList(value, int.MaxValue, out values);

    private static bool TryParseNumberList(string? value, int maximumValues,
        out IReadOnlyList<double> values) =>
        TryParseNumberList(value, maximumValues, out values, out _);

    private static bool TryParseNumberList(string? value, int maximumValues,
        out IReadOnlyList<double> values, out bool limitExceeded) {
        var result = new List<double>(Math.Min(maximumValues, 16));
        values = result;
        limitExceeded = false;
        if (maximumValues <= 0 || string.IsNullOrWhiteSpace(value)) return false;
        int index = 0;
        while (index < value!.Length) {
            while (index < value.Length && (char.IsWhiteSpace(value[index]) || value[index] == ',')) index++;
            if (index >= value.Length) break;
            if (result.Count >= maximumValues) {
                limitExceeded = true;
                return false;
            }
            int start = index;
            while (index < value.Length && !char.IsWhiteSpace(value[index]) && value[index] != ',') {
                index++;
                if (index - start > 128) {
                    limitExceeded = true;
                    return false;
                }
            }
            int length = index - start;
            if (length <= 0
                || !double.TryParse(value.Substring(start, length), NumberStyles.Float,
                    CultureInfo.InvariantCulture, out double number)
                || double.IsNaN(number)
                || double.IsInfinity(number)) return false;
            result.Add(number);
        }
        return result.Count > 0;
    }

    private struct SvgPaintContext {
        internal OfficeColor Color;
        internal OfficeColor? Fill;
        internal OfficeLinearGradient? FillGradient;
        internal OfficeRadialGradient? FillRadialGradient;
        internal SvgGradientDefinition? FillDeferredGradient;
        internal OfficeColor? Stroke;
        internal OfficeLinearGradient? StrokeGradient;
        internal OfficeRadialGradient? StrokeRadialGradient;
        internal SvgGradientDefinition? StrokeDeferredGradient;
        internal double StrokeWidth;
        internal double Opacity;
        internal double FillOpacity;
        internal double StrokeOpacity;
        internal OfficeStrokeDashStyle DashStyle;
        internal OfficeStrokeLineCap LineCap;
        internal OfficeStrokeLineJoin LineJoin;
        internal OfficeFillRule FillRule;
        internal string FontFamily;
        internal double FontSize;
        internal OfficeFontStyle FontStyle;
        internal string TextAnchor;
        internal SvgDominantBaseline DominantBaseline;
        internal bool Visible;

        internal void SetFill(SvgResolvedPaint paint) {
            Fill = paint.Color;
            FillGradient = paint.LinearGradient;
            FillRadialGradient = paint.RadialGradient;
            FillDeferredGradient = paint.DeferredGradient;
        }

        internal void SetStroke(SvgResolvedPaint paint) {
            Stroke = paint.Color;
            StrokeGradient = paint.LinearGradient;
            StrokeRadialGradient = paint.RadialGradient;
            StrokeDeferredGradient = paint.DeferredGradient;
        }

        internal static SvgPaintContext Default => new SvgPaintContext {
            Color = OfficeColor.Black,
            Fill = OfficeColor.Black,
            Stroke = null,
            StrokeWidth = 1D,
            Opacity = 1D,
            FillOpacity = 1D,
            StrokeOpacity = 1D,
            DashStyle = OfficeStrokeDashStyle.Solid,
            LineCap = OfficeStrokeLineCap.Butt,
            LineJoin = OfficeStrokeLineJoin.Miter,
            FillRule = OfficeFillRule.NonZero,
            FontFamily = "Arial",
            FontSize = 16D,
            FontStyle = OfficeFontStyle.Regular,
            TextAnchor = "start",
            DominantBaseline = SvgDominantBaseline.Alphabetic,
            Visible = true
        };
    }

    private enum SvgDominantBaseline {
        Alphabetic,
        Hanging,
        Middle,
        TextAfterEdge
    }
}
