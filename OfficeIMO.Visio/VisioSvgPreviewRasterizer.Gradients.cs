using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;

namespace OfficeIMO.Visio {
    internal static partial class VisioSvgPreviewRasterizer {
        private static bool TryResolveFillGradient(string? rawFill, SvgRenderContext context, double opacity, OfficeColor currentColor, out OfficeLinearGradient? linear, out OfficeRadialGradient? radial) {
            linear = null;
            radial = null;
            if (!TryReadUrlId(rawFill, out string? id) || id == null || !context.TryGetDefinition(id, out XElement? definition) || definition == null) {
                return false;
            }

            string name = definition.Name.LocalName;
            if (string.Equals(name, "linearGradient", StringComparison.OrdinalIgnoreCase)) {
                linear = TryCreateLinearGradient(definition, context, opacity, currentColor, out OfficeLinearGradient? gradient) ? gradient : null;
            } else if (string.Equals(name, "radialGradient", StringComparison.OrdinalIgnoreCase)) {
                radial = TryCreateRadialGradient(definition, context, opacity, currentColor, out OfficeRadialGradient? gradient) ? gradient : null;
            }

            return linear != null || radial != null;
        }

        private static bool TryCreateLinearGradient(XElement definition, SvgRenderContext context, double opacity, OfficeColor currentColor, out OfficeLinearGradient? gradient) {
            gradient = null;
            if (!TryReadGradientStops(definition, context, opacity, currentColor, out IReadOnlyList<OfficeGradientStop> stops)) {
                return false;
            }

            bool userSpace = IsUserSpaceGradient(definition, context);
            double x1 = ReadGradientUnit(ReadInheritedGradientAttribute(definition, context, "x1"), 0D);
            double y1 = ReadGradientUnit(ReadInheritedGradientAttribute(definition, context, "y1"), 0D);
            double x2 = ReadGradientUnit(ReadInheritedGradientAttribute(definition, context, "x2"), userSpace ? x1 + 1D : 1D);
            double y2 = ReadGradientUnit(ReadInheritedGradientAttribute(definition, context, "y2"), 0D);
            SvgTransform gradientTransform = ReadInheritedGradientTransform(definition, context);
            if (userSpace) {
                OfficePoint start = gradientTransform.Apply(x1, y1);
                OfficePoint end = gradientTransform.Apply(x2, y2);
                x1 = start.X;
                y1 = start.Y;
                x2 = end.X;
                y2 = end.Y;
            }

            if (userSpace && context.CurrentPaintBounds.HasValue) {
                SvgPaintBounds bounds = context.CurrentPaintBounds.Value;
                if (bounds.HasArea) {
                    x1 = NormalizeUserSpaceGradientCoordinate(x1, bounds.Left, bounds.Width);
                    y1 = NormalizeUserSpaceGradientCoordinate(y1, bounds.Top, bounds.Height);
                    x2 = NormalizeUserSpaceGradientCoordinate(x2, bounds.Left, bounds.Width);
                    y2 = NormalizeUserSpaceGradientCoordinate(y2, bounds.Top, bounds.Height);
                }
            }

            if (!userSpace) {
                OfficePoint start = gradientTransform.Apply(x1, y1);
                OfficePoint end = gradientTransform.Apply(x2, y2);
                x1 = start.X;
                y1 = start.Y;
                x2 = end.X;
                y2 = end.Y;
            }

            if (x1.Equals(x2) && y1.Equals(y2)) {
                x2 = x1 < 1D ? 1D : 0D;
            }

            if (!TryClipLinearGradientToUnitBox(x1, y1, x2, y2, stops, out OfficePoint clippedStart, out OfficePoint clippedEnd, out IReadOnlyList<OfficeGradientStop> clippedStops)) {
                return false;
            }

            gradient = new OfficeLinearGradient(clippedStart.X, clippedStart.Y, clippedEnd.X, clippedEnd.Y, clippedStops);
            return true;
        }

        private static bool TryCreateRadialGradient(XElement definition, SvgRenderContext context, double opacity, OfficeColor currentColor, out OfficeRadialGradient? gradient) {
            gradient = null;
            if (!TryReadGradientStops(definition, context, opacity, currentColor, out IReadOnlyList<OfficeGradientStop> stops)) {
                return false;
            }

            bool userSpace = IsUserSpaceGradient(definition, context);
            double cx = ReadGradientUnit(ReadInheritedGradientAttribute(definition, context, "cx"), 0.5D);
            double cy = ReadGradientUnit(ReadInheritedGradientAttribute(definition, context, "cy"), 0.5D);
            double r = ReadGradientUnit(ReadInheritedGradientAttribute(definition, context, "r"), 0.5D);
            double fx = ReadGradientUnit(ReadInheritedGradientAttribute(definition, context, "fx"), cx);
            double fy = ReadGradientUnit(ReadInheritedGradientAttribute(definition, context, "fy"), cy);
            double fr = ReadGradientUnit(ReadInheritedGradientAttribute(definition, context, "fr"), 0D);
            SvgTransform gradientTransform = ReadInheritedGradientTransform(definition, context);
            double radiusScale = Math.Max(gradientTransform.ScaleX, gradientTransform.ScaleY);
            if (userSpace) {
                OfficePoint center = gradientTransform.Apply(cx, cy);
                OfficePoint focus = gradientTransform.Apply(fx, fy);
                cx = center.X;
                cy = center.Y;
                fx = focus.X;
                fy = focus.Y;
                r *= radiusScale;
                fr *= radiusScale;
            }

            if (userSpace && context.CurrentPaintBounds.HasValue) {
                SvgPaintBounds bounds = context.CurrentPaintBounds.Value;
                if (bounds.HasArea) {
                    cx = NormalizeUserSpaceGradientCoordinate(cx, bounds.Left, bounds.Width);
                    cy = NormalizeUserSpaceGradientCoordinate(cy, bounds.Top, bounds.Height);
                    fx = NormalizeUserSpaceGradientCoordinate(fx, bounds.Left, bounds.Width);
                    fy = NormalizeUserSpaceGradientCoordinate(fy, bounds.Top, bounds.Height);
                    r = NormalizeUserSpaceGradientRadius(r, bounds);
                    fr = NormalizeUserSpaceGradientRadius(fr, bounds);
                }
            }

            if (!userSpace) {
                OfficePoint center = gradientTransform.Apply(cx, cy);
                OfficePoint focus = gradientTransform.Apply(fx, fy);
                cx = center.X;
                cy = center.Y;
                fx = focus.X;
                fy = focus.Y;
                r *= radiusScale;
                fr *= radiusScale;
            }

            if (r.Equals(fr) && cx.Equals(fx) && cy.Equals(fy)) {
                r = Math.Min(1D, fr + 0.5D);
            }

            gradient = new OfficeRadialGradient(ClampUnit(fx), ClampUnit(fy), Math.Max(0D, fr), ClampUnit(cx), ClampUnit(cy), Math.Max(0D, r), stops);
            return true;
        }

        private static bool TryReadGradientStops(XElement definition, SvgRenderContext context, double opacity, OfficeColor currentColor, out IReadOnlyList<OfficeGradientStop> stops) {
            List<OfficeGradientStop> parsedStops = new();
            IEnumerable<XElement> stopElements = definition.Elements().Where(element => string.Equals(element.Name.LocalName, "stop", StringComparison.OrdinalIgnoreCase));
            if (!stopElements.Any() && TryReadUrlId(ReadHref(definition), out string? hrefId) && hrefId != null && context.TryGetDefinition(hrefId, out XElement? inherited) && inherited != null) {
                stopElements = inherited.Elements().Where(element => string.Equals(element.Name.LocalName, "stop", StringComparison.OrdinalIgnoreCase));
            }

            foreach (XElement stopElement in stopElements) {
                Dictionary<string, string> stopStyle = context.StyleSheet.CreateStyle(stopElement);
                double offset = ReadGradientUnit(stopElement.Attribute("offset")?.Value, 0D);
                string? colorValue = stopStyle.TryGetValue("stop-color", out string? styleColor)
                    ? styleColor
                    : stopElement.Attribute("stop-color")?.Value;
                if (string.IsNullOrWhiteSpace(colorValue)) {
                    colorValue = "black";
                }

                double stopOpacity = ReadOpacity(
                    stopStyle.TryGetValue("stop-opacity", out string? styleOpacity)
                        ? styleOpacity
                        : stopElement.Attribute("stop-opacity")?.Value,
                    1D);
                if (TryReadGradientColor(colorValue, opacity * stopOpacity, currentColor, out OfficeColor color)) {
                    parsedStops.Add(new OfficeGradientStop(ClampUnit(offset), color));
                }
            }

            stops = NormalizeStops(parsedStops);
            return stops.Count >= 2;
        }

        private static IReadOnlyList<OfficeGradientStop> NormalizeStops(List<OfficeGradientStop> stops) {
            if (stops.Count == 0) {
                return Array.Empty<OfficeGradientStop>();
            }

            stops.Sort((left, right) => left.Offset.CompareTo(right.Offset));
            List<OfficeGradientStop> normalized = new();
            for (int i = 0; i < stops.Count; i++) {
                if (normalized.Count > 0 && normalized[normalized.Count - 1].Offset.Equals(stops[i].Offset)) {
                    normalized[normalized.Count - 1] = stops[i];
                } else {
                    normalized.Add(stops[i]);
                }
            }

            if (normalized[0].Offset > 0D) {
                normalized.Insert(0, new OfficeGradientStop(0D, normalized[0].Color));
            }

            if (normalized[normalized.Count - 1].Offset < 1D) {
                normalized.Add(new OfficeGradientStop(1D, normalized[normalized.Count - 1].Color));
            }

            return normalized;
        }

        private static bool TryReadUrlId(string? raw, out string? id) {
            id = null;
            if (string.IsNullOrWhiteSpace(raw)) {
                return false;
            }

            string value = raw!.Trim();
            int start = value.IndexOf("url(", StringComparison.OrdinalIgnoreCase);
            if (start < 0) {
                return false;
            }

            int open = start + 4;
            int close = value.IndexOf(')', open);
            if (close <= open) {
                return false;
            }

            string reference = value.Substring(open, close - open).Trim().Trim('\'', '"');
            if (reference.Length < 2 || reference[0] != '#') {
                return false;
            }

            id = reference.Substring(1);
            return id.Length > 0;
        }

        private static bool TryReadGradientColor(string? raw, double opacity, OfficeColor currentColor, out OfficeColor color) {
            color = OfficeColor.Transparent;
            if (string.IsNullOrWhiteSpace(raw)) {
                return false;
            }

            string value = raw!.Trim();
            if (string.Equals(value, "none", StringComparison.OrdinalIgnoreCase)) {
                color = OfficeColor.Transparent;
                return true;
            }

            if (string.Equals(value, "currentColor", StringComparison.OrdinalIgnoreCase)) {
                color = ApplyAlpha(currentColor, opacity);
                return true;
            }

            if (TryParseRgbColor(value, out OfficeColor rgbColor)) {
                color = ApplyAlpha(rgbColor, opacity);
                return true;
            }

            if (OfficeColor.TryParse(value, out OfficeColor parsed)) {
                color = ApplyAlpha(parsed, opacity);
                return true;
            }

            return false;
        }

        private static Dictionary<string, string> ReadDeclarations(string? raw) {
            Dictionary<string, string> declarations = new(StringComparer.OrdinalIgnoreCase);
            if (string.IsNullOrWhiteSpace(raw)) {
                return declarations;
            }

            string[] parts = raw!.Split(';');
            for (int i = 0; i < parts.Length; i++) {
                int separator = parts[i].IndexOf(':');
                if (separator <= 0) {
                    continue;
                }

                declarations[parts[i].Substring(0, separator).Trim()] = parts[i].Substring(separator + 1).Trim();
            }

            return declarations;
        }

        private static double ReadGradientUnit(string? raw, double fallback) {
            if (string.IsNullOrWhiteSpace(raw)) {
                return fallback;
            }

            string value = raw!.Trim();
            bool percent = value.EndsWith("%", StringComparison.Ordinal);
            if (percent) {
                value = value.Substring(0, value.Length - 1);
            }

            if (!double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)) {
                return fallback;
            }

            return percent ? parsed / 100D : parsed;
        }

        private static string? ReadInheritedGradientAttribute(XElement definition, SvgRenderContext context, string attributeName) {
            string? value = definition.Attribute(attributeName)?.Value;
            if (!string.IsNullOrWhiteSpace(value)) {
                return value;
            }

            return ReadInheritedGradientAttribute(definition, context, attributeName, new HashSet<string>(StringComparer.Ordinal));
        }

        private static string? ReadInheritedGradientAttribute(XElement definition, SvgRenderContext context, string attributeName, HashSet<string> visited) {
            if (!TryReadUrlId(ReadHref(definition), out string? hrefId) ||
                hrefId == null ||
                !visited.Add(hrefId) ||
                !context.TryGetDefinition(hrefId, out XElement? inherited) ||
                inherited == null) {
                return null;
            }

            string? value = inherited.Attribute(attributeName)?.Value;
            return !string.IsNullOrWhiteSpace(value)
                ? value
                : ReadInheritedGradientAttribute(inherited, context, attributeName, visited);
        }

        private static bool IsUserSpaceGradient(XElement definition, SvgRenderContext context) =>
            string.Equals(ReadInheritedGradientAttribute(definition, context, "gradientUnits"), "userSpaceOnUse", StringComparison.OrdinalIgnoreCase);

        private static SvgTransform ReadInheritedGradientTransform(XElement definition, SvgRenderContext context) {
            string? value = ReadInheritedGradientAttribute(definition, context, "gradientTransform");
            return string.IsNullOrWhiteSpace(value) ? SvgTransform.Identity : ReadTransform(value);
        }

        private static double NormalizeUserSpaceGradientCoordinate(double value, double origin, double length) =>
            length > 0D ? (value - origin) / length : value;

        private static double NormalizeUserSpaceGradientRadius(double value, SvgPaintBounds bounds) {
            double length = Math.Max(bounds.Width, bounds.Height);
            return length > 0D ? value / length : value;
        }

        private static double ReadOpacity(string? raw, double fallback) {
            if (string.IsNullOrWhiteSpace(raw) || !double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)) {
                return fallback;
            }

            return ClampUnit(parsed);
        }

        private static OfficeColor ApplyAlpha(OfficeColor color, double opacity) =>
            OfficeColor.FromRgba(color.R, color.G, color.B, (byte)Math.Round(color.A * ClampUnit(opacity)));

        private static byte ToByte(double value) => (byte)Math.Max(0D, Math.Min(255D, Math.Round(value)));

        private static bool TryClipLinearGradientToUnitBox(
            double x1,
            double y1,
            double x2,
            double y2,
            IReadOnlyList<OfficeGradientStop> stops,
            out OfficePoint clippedStart,
            out OfficePoint clippedEnd,
            out IReadOnlyList<OfficeGradientStop> clippedStops) {
            clippedStart = default;
            clippedEnd = default;
            clippedStops = Array.Empty<OfficeGradientStop>();
            double dx = x2 - x1;
            double dy = y2 - y1;
            double t0 = 0D;
            double t1 = 1D;
            if (!ClipLinearGradientEdge(-dx, x1, ref t0, ref t1) ||
                !ClipLinearGradientEdge(dx, 1D - x1, ref t0, ref t1) ||
                !ClipLinearGradientEdge(-dy, y1, ref t0, ref t1) ||
                !ClipLinearGradientEdge(dy, 1D - y1, ref t0, ref t1) ||
                t1 <= t0) {
                return false;
            }

            clippedStart = new OfficePoint(ClampUnit(x1 + dx * t0), ClampUnit(y1 + dy * t0));
            clippedEnd = new OfficePoint(ClampUnit(x1 + dx * t1), ClampUnit(y1 + dy * t1));
            var adjusted = new List<OfficeGradientStop>();
            adjusted.Add(new OfficeGradientStop(0D, InterpolateGradientColor(stops, t0)));
            for (int i = 0; i < stops.Count; i++) {
                double offset = stops[i].Offset;
                if (offset <= t0 || offset >= t1) {
                    continue;
                }

                adjusted.Add(new OfficeGradientStop((offset - t0) / (t1 - t0), stops[i].Color));
            }

            adjusted.Add(new OfficeGradientStop(1D, InterpolateGradientColor(stops, t1)));
            clippedStops = NormalizeStops(adjusted);
            return clippedStops.Count >= 2 && !(clippedStart.X.Equals(clippedEnd.X) && clippedStart.Y.Equals(clippedEnd.Y));
        }

        private static bool ClipLinearGradientEdge(double p, double q, ref double t0, ref double t1) {
            if (Math.Abs(p) < 0.0000001D) {
                return q >= 0D;
            }

            double r = q / p;
            if (p < 0D) {
                if (r > t1) {
                    return false;
                }

                if (r > t0) {
                    t0 = r;
                }
            } else {
                if (r < t0) {
                    return false;
                }

                if (r < t1) {
                    t1 = r;
                }
            }

            return true;
        }

        private static OfficeColor InterpolateGradientColor(IReadOnlyList<OfficeGradientStop> stops, double offset) {
            if (offset <= stops[0].Offset) {
                return stops[0].Color;
            }

            for (int i = 1; i < stops.Count; i++) {
                if (offset > stops[i].Offset) {
                    continue;
                }

                OfficeGradientStop left = stops[i - 1];
                OfficeGradientStop right = stops[i];
                double span = right.Offset - left.Offset;
                double t = span <= 0D ? 0D : (offset - left.Offset) / span;
                return OfficeColor.FromRgba(
                    ToByte(left.Color.R + (right.Color.R - left.Color.R) * t),
                    ToByte(left.Color.G + (right.Color.G - left.Color.G) * t),
                    ToByte(left.Color.B + (right.Color.B - left.Color.B) * t),
                    ToByte(left.Color.A + (right.Color.A - left.Color.A) * t));
            }

            return stops[stops.Count - 1].Color;
        }

        private static double ClampUnit(double value) => value < 0D ? 0D : value > 1D ? 1D : value;
    }
}
