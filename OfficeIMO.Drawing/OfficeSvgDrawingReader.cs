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
    private const int MaximumElements = 10000;

    /// <summary>Attempts to interpret supported SVG vector primitives as a shared drawing.</summary>
    public static bool TryRead(byte[]? bytes, out OfficeDrawing? drawing) =>
        TryRead(bytes, out drawing, out _);

    /// <summary>
    /// Attempts to interpret supported SVG vector primitives as a shared drawing and reports the
    /// number of elements or declarations that required omission or fallback.
    /// </summary>
    public static bool TryRead(byte[]? bytes, out OfficeDrawing? drawing, out int unsupportedFeatureCount) {
        drawing = null;
        unsupportedFeatureCount = 0;
        if (bytes == null || bytes.Length == 0 || bytes.Length > MaximumInputBytes) return false;

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
            if (root.Descendants().Take(MaximumElements + 1).Count() > MaximumElements) return false;
            if (!TryResolveViewport(bytes, root, out double viewX, out double viewY, out double width, out double height)) return false;

            var result = new OfficeDrawing(width, height);
            int visited = 0;
            SvgDefinitionRegistry definitions = SvgDefinitionRegistry.Create(root);
            var paintServers = new SvgPaintServerRegistry(definitions);
            var references = new SvgElementReferenceRegistry(definitions);
            var context = ResolvePaintContext(root, SvgPaintContext.Default, paintServers, ref unsupportedFeatureCount);
            OfficeTransform rootTransform = ResolveTransform(root, OfficeTransform.Identity, viewX, viewY, ref unsupportedFeatureCount);
            AddChildren(root, result, context, paintServers, references, rootTransform, viewX, viewY, ref visited, ref unsupportedFeatureCount);
            if (visited > MaximumElements) return false;
            drawing = result;
            return true;
        } catch (XmlException) {
            return false;
        } catch (InvalidOperationException) {
            return false;
        } catch (ArgumentException) {
            return false;
        }
    }

    private static bool TryResolveViewport(byte[] bytes, XElement root, out double viewX, out double viewY, out double width, out double height) {
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
            return true;
        }

        if (!OfficeImageReader.TryIdentify(bytes, ".svg", out OfficeImageInfo info) || info.Width <= 0 || info.Height <= 0) return false;
        width = info.Width * 96D / Math.Max(1D, info.DpiX);
        height = info.Height * 96D / Math.Max(1D, info.DpiY);
        return width > 0D && height > 0D;
    }

    private static void AddChildren(
        XElement parent,
        OfficeDrawing drawing,
        SvgPaintContext inherited,
        SvgPaintServerRegistry paintServers,
        SvgElementReferenceRegistry references,
        OfficeTransform inheritedTransform,
        double viewX,
        double viewY,
        ref int visited,
        ref int unsupported) {
        foreach (XElement element in parent.Elements()) {
            AddElement(element, drawing, inherited, paintServers, references, inheritedTransform, viewX, viewY, ref visited, ref unsupported);
            if (visited > MaximumElements) return;
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
        ref int visited,
        ref int unsupported) {
        visited++;
        if (visited > MaximumElements) return;
        string name = element.Name.LocalName.ToLowerInvariant();
        if (name is "title" or "desc" or "metadata" or "lineargradient" or "radialgradient" or "stop") return;
        if (name == "defs") return;

        SvgPaintContext style = ResolvePaintContext(element, inherited, paintServers, ref unsupported);
        if (!style.Visible) return;
        OfficeTransform transform = ResolveTransform(element, inheritedTransform, viewX, viewY, ref unsupported);
        if (name is "g" or "svg" or "a" or "switch") {
            AddChildren(element, drawing, style, paintServers, references, transform, viewX, viewY, ref visited, ref unsupported);
            return;
        }
        if (name == "use") {
            AddReferencedElement(element, drawing, style, paintServers, references, transform, viewX, viewY, ref visited, ref unsupported);
            return;
        }
        if (name == "text") {
            AddText(element, drawing, style, paintServers, transform, viewX, viewY, ref unsupported);
            return;
        }

        OfficeDrawingShape? shape = name switch {
            "rect" => CreateRectangle(element, style, viewX, viewY, ref unsupported),
            "circle" => CreateCircle(element, style, viewX, viewY),
            "ellipse" => CreateEllipse(element, style, viewX, viewY),
            "line" => CreateLine(element, style, viewX, viewY),
            "polygon" => CreatePolygon(element, style, viewX, viewY, close: true),
            "polyline" => CreatePolygon(element, style, viewX, viewY, close: false),
            "path" => CreatePath(element, style, viewX, viewY),
            _ => null
        };
        if (shape == null) {
            unsupported++;
            return;
        }

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
        return normalized.Then(inherited);
    }

    private static void ApplyTransform(OfficeDrawingShape drawingShape, OfficeTransform transform) {
        if (transform == OfficeTransform.Identity) return;
        OfficeTransform local = OfficeTransform.Translate(drawingShape.X, drawingShape.Y)
            .Then(transform)
            .Then(OfficeTransform.Translate(-drawingShape.X, -drawingShape.Y));
        OfficeShape shape = drawingShape.Shape;
        shape.Transform = shape.Transform.HasValue ? shape.Transform.Value.Then(local) : local;
    }

    private static OfficeDrawingShape? CreateRectangle(XElement element, SvgPaintContext style, double viewX, double viewY, ref int unsupported) {
        if (!TryLength(element, "width", out double width) || !TryLength(element, "height", out double height) || width <= 0D || height <= 0D) return null;
        double x = ReadLength(element, "x") - viewX;
        double y = ReadLength(element, "y") - viewY;
        double rx = ReadLength(element, "rx");
        double ry = ReadLength(element, "ry");
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

    private static OfficeDrawingShape? CreateCircle(XElement element, SvgPaintContext style, double viewX, double viewY) {
        if (!TryLength(element, "r", out double radius) || radius <= 0D) return null;
        double x = ReadLength(element, "cx") - radius - viewX;
        double y = ReadLength(element, "cy") - radius - viewY;
        OfficeShape shape = OfficeShape.Ellipse(radius * 2D, radius * 2D);
        ApplyPaint(shape, style);
        return new OfficeDrawingShape(shape, x, y);
    }

    private static OfficeDrawingShape? CreateEllipse(XElement element, SvgPaintContext style, double viewX, double viewY) {
        if (!TryLength(element, "rx", out double radiusX) || !TryLength(element, "ry", out double radiusY) || radiusX <= 0D || radiusY <= 0D) return null;
        double x = ReadLength(element, "cx") - radiusX - viewX;
        double y = ReadLength(element, "cy") - radiusY - viewY;
        OfficeShape shape = OfficeShape.Ellipse(radiusX * 2D, radiusY * 2D);
        ApplyPaint(shape, style);
        return new OfficeDrawingShape(shape, x, y);
    }

    private static OfficeDrawingShape? CreateLine(XElement element, SvgPaintContext style, double viewX, double viewY) {
        double x1 = ReadLength(element, "x1") - viewX;
        double y1 = ReadLength(element, "y1") - viewY;
        double x2 = ReadLength(element, "x2") - viewX;
        double y2 = ReadLength(element, "y2") - viewY;
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

    private static OfficeDrawingShape? CreatePolygon(XElement element, SvgPaintContext style, double viewX, double viewY, bool close) {
        if (!TryParseNumberList(element.Attribute("points")?.Value, out IReadOnlyList<double> values) || values.Count < 4 || values.Count % 2 != 0) return null;
        var points = new List<OfficePoint>(values.Count / 2);
        for (int index = 0; index < values.Count; index += 2) points.Add(new OfficePoint(values[index] - viewX, values[index + 1] - viewY));
        double minX = points.Min(point => point.X);
        double minY = points.Min(point => point.Y);
        OfficeShape shape;
        if (close) {
            if (points.Count < 3) return null;
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
        ApplyPaint(shape, style);
        return new OfficeDrawingShape(shape, minX, minY);
    }

    private static OfficeDrawingShape? CreatePath(XElement element, SvgPaintContext style, double viewX, double viewY) {
        if (!OfficeSvgPathDataParser.TryParse(element.Attribute("d")?.Value, out IReadOnlyList<OfficePathCommand> parsed)) return null;
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
        ApplyProperty("display", element.Attribute("display")?.Value, paintServers, ref result, ref unsupported);
        ApplyProperty("visibility", element.Attribute("visibility")?.Value, paintServers, ref result, ref unsupported);
        string? declarations = element.Attribute("style")?.Value;
        if (!string.IsNullOrWhiteSpace(declarations)) {
            foreach (string declaration in declarations!.Split(';')) {
                int colon = declaration.IndexOf(':');
                if (colon <= 0) continue;
                ApplyProperty(declaration.Substring(0, colon).Trim(), declaration.Substring(colon + 1).Trim(), paintServers, ref result, ref unsupported);
            }
        }
        return result;
    }

    private static void ApplyProperty(string name, string? value, SvgPaintServerRegistry paintServers, ref SvgPaintContext style, ref int unsupported) {
        if (string.IsNullOrWhiteSpace(value)) return;
        string normalized = value!.Trim();
        switch (name.Trim().ToLowerInvariant()) {
            case "fill":
                if (!TryPaint(normalized, paintServers, out SvgResolvedPaint fill)) {
                    unsupported++;
                    if (normalized.StartsWith("url(", StringComparison.OrdinalIgnoreCase)) style.SetFill(default);
                }
                else style.SetFill(fill);
                break;
            case "stroke":
                if (!TryPaint(normalized, paintServers, out SvgResolvedPaint stroke)) {
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

    private static double ReadLength(XElement element, string name) => TryLength(element, name, out double value) ? value : 0D;
    private static bool TryLength(XElement element, string name, out double value) => TrySvgLength(element.Attribute(name)?.Value, out value);

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

    private static bool TryParseNumberList(string? value, out IReadOnlyList<double> values) {
        var result = new List<double>();
        values = result;
        if (string.IsNullOrWhiteSpace(value)) return false;
        string normalized = value!.Replace(',', ' ');
        foreach (string part in normalized.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)) {
            if (!double.TryParse(part, NumberStyles.Float, CultureInfo.InvariantCulture, out double number)
                || double.IsNaN(number)
                || double.IsInfinity(number)) return false;
            result.Add(number);
        }
        return result.Count > 0;
    }

    private struct SvgPaintContext {
        internal OfficeColor? Fill;
        internal OfficeLinearGradient? FillGradient;
        internal OfficeRadialGradient? FillRadialGradient;
        internal OfficeColor? Stroke;
        internal OfficeLinearGradient? StrokeGradient;
        internal OfficeRadialGradient? StrokeRadialGradient;
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
        internal bool Visible;

        internal void SetFill(SvgResolvedPaint paint) {
            Fill = paint.Color;
            FillGradient = paint.LinearGradient;
            FillRadialGradient = paint.RadialGradient;
        }

        internal void SetStroke(SvgResolvedPaint paint) {
            Stroke = paint.Color;
            StrokeGradient = paint.LinearGradient;
            StrokeRadialGradient = paint.RadialGradient;
        }

        internal static SvgPaintContext Default => new SvgPaintContext {
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
            Visible = true
        };
    }
}
