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
public static class OfficeSvgDrawingReader {
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
            if (!TryResolveViewport(bytes, root, out double viewX, out double viewY, out double width, out double height)) return false;

            var result = new OfficeDrawing(width, height);
            int visited = 0;
            var context = SvgPaintContext.Default;
            AddChildren(root, result, context, viewX, viewY, ref visited, ref unsupportedFeatureCount);
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
        double viewX,
        double viewY,
        ref int visited,
        ref int unsupported) {
        foreach (XElement element in parent.Elements()) {
            visited++;
            if (visited > MaximumElements) return;
            string name = element.Name.LocalName.ToLowerInvariant();
            if (name is "title" or "desc" or "metadata") continue;
            if (name == "defs") {
                if (element.Elements().Any()) unsupported++;
                continue;
            }

            SvgPaintContext style = ResolvePaintContext(element, inherited, ref unsupported);
            if (!style.Visible) continue;
            if (element.Attribute("transform") != null) unsupported++;
            if (name is "g" or "svg" or "a" or "switch") {
                AddChildren(element, drawing, style, viewX, viewY, ref visited, ref unsupported);
                continue;
            }

            OfficeDrawingShape? shape = name switch {
                "rect" => CreateRectangle(element, style, viewX, viewY, ref unsupported),
                "circle" => CreateCircle(element, style, viewX, viewY),
                "ellipse" => CreateEllipse(element, style, viewX, viewY),
                "line" => CreateLine(element, style, viewX, viewY),
                "polygon" => CreatePolygon(element, style, viewX, viewY, close: true),
                "polyline" => CreatePolygon(element, style, viewX, viewY, close: false),
                _ => null
            };
            if (shape == null) {
                unsupported++;
                continue;
            }

            try {
                drawing.AddShape(shape.Shape, shape.X, shape.Y);
            } catch (ArgumentOutOfRangeException) {
                unsupported++;
            }
        }
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

    private static SvgPaintContext ResolvePaintContext(XElement element, SvgPaintContext inherited, ref int unsupported) {
        SvgPaintContext result = inherited;
        ApplyProperty("fill", element.Attribute("fill")?.Value, ref result, ref unsupported);
        ApplyProperty("stroke", element.Attribute("stroke")?.Value, ref result, ref unsupported);
        ApplyProperty("stroke-width", element.Attribute("stroke-width")?.Value, ref result, ref unsupported);
        ApplyProperty("opacity", element.Attribute("opacity")?.Value, ref result, ref unsupported);
        ApplyProperty("fill-opacity", element.Attribute("fill-opacity")?.Value, ref result, ref unsupported);
        ApplyProperty("stroke-opacity", element.Attribute("stroke-opacity")?.Value, ref result, ref unsupported);
        ApplyProperty("stroke-dasharray", element.Attribute("stroke-dasharray")?.Value, ref result, ref unsupported);
        ApplyProperty("stroke-linecap", element.Attribute("stroke-linecap")?.Value, ref result, ref unsupported);
        ApplyProperty("stroke-linejoin", element.Attribute("stroke-linejoin")?.Value, ref result, ref unsupported);
        ApplyProperty("fill-rule", element.Attribute("fill-rule")?.Value, ref result, ref unsupported);
        ApplyProperty("display", element.Attribute("display")?.Value, ref result, ref unsupported);
        ApplyProperty("visibility", element.Attribute("visibility")?.Value, ref result, ref unsupported);
        string? declarations = element.Attribute("style")?.Value;
        if (!string.IsNullOrWhiteSpace(declarations)) {
            foreach (string declaration in declarations!.Split(';')) {
                int colon = declaration.IndexOf(':');
                if (colon <= 0) continue;
                ApplyProperty(declaration.Substring(0, colon).Trim(), declaration.Substring(colon + 1).Trim(), ref result, ref unsupported);
            }
        }
        return result;
    }

    private static void ApplyProperty(string name, string? value, ref SvgPaintContext style, ref int unsupported) {
        if (string.IsNullOrWhiteSpace(value)) return;
        string normalized = value!.Trim();
        switch (name.Trim().ToLowerInvariant()) {
            case "fill":
                if (!TryPaint(normalized, out OfficeColor? fill)) unsupported++;
                else style.Fill = fill;
                break;
            case "stroke":
                if (!TryPaint(normalized, out OfficeColor? stroke)) unsupported++;
                else style.Stroke = stroke;
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

    private static bool TryPaint(string value, out OfficeColor? color) {
        color = null;
        if (value.Equals("none", StringComparison.OrdinalIgnoreCase)) return true;
        if (value.Equals("currentcolor", StringComparison.OrdinalIgnoreCase) || value.StartsWith("url(", StringComparison.OrdinalIgnoreCase)) return false;
        if (!OfficeColor.TryParse(value, out OfficeColor parsed)) return false;
        color = parsed;
        return true;
    }

    private static void ApplyPaint(OfficeShape shape, SvgPaintContext style) {
        shape.FillColor = style.Fill;
        shape.StrokeColor = style.Stroke;
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
        internal OfficeColor? Stroke;
        internal double StrokeWidth;
        internal double Opacity;
        internal double FillOpacity;
        internal double StrokeOpacity;
        internal OfficeStrokeDashStyle DashStyle;
        internal OfficeStrokeLineCap LineCap;
        internal OfficeStrokeLineJoin LineJoin;
        internal OfficeFillRule FillRule;
        internal bool Visible;

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
            Visible = true
        };
    }
}
