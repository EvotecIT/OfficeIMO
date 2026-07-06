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

            bool userSpace = IsUserSpaceGradient(definition);
            double x1 = ReadGradientUnit(definition.Attribute("x1")?.Value, 0D);
            double y1 = ReadGradientUnit(definition.Attribute("y1")?.Value, 0D);
            double x2 = ReadGradientUnit(definition.Attribute("x2")?.Value, userSpace ? x1 + 1D : 1D);
            double y2 = ReadGradientUnit(definition.Attribute("y2")?.Value, 0D);
            if (userSpace && context.CurrentPaintBounds.HasValue) {
                SvgPaintBounds bounds = context.CurrentPaintBounds.Value;
                if (bounds.HasArea) {
                    x1 = NormalizeUserSpaceGradientCoordinate(x1, bounds.Left, bounds.Width);
                    y1 = NormalizeUserSpaceGradientCoordinate(y1, bounds.Top, bounds.Height);
                    x2 = NormalizeUserSpaceGradientCoordinate(x2, bounds.Left, bounds.Width);
                    y2 = NormalizeUserSpaceGradientCoordinate(y2, bounds.Top, bounds.Height);
                }
            }

            if (x1.Equals(x2) && y1.Equals(y2)) {
                x2 = x1 < 1D ? 1D : 0D;
            }

            gradient = new OfficeLinearGradient(ClampUnit(x1), ClampUnit(y1), ClampUnit(x2), ClampUnit(y2), stops);
            return true;
        }

        private static bool TryCreateRadialGradient(XElement definition, SvgRenderContext context, double opacity, OfficeColor currentColor, out OfficeRadialGradient? gradient) {
            gradient = null;
            if (!TryReadGradientStops(definition, context, opacity, currentColor, out IReadOnlyList<OfficeGradientStop> stops)) {
                return false;
            }

            bool userSpace = IsUserSpaceGradient(definition);
            double cx = ReadGradientUnit(definition.Attribute("cx")?.Value, 0.5D);
            double cy = ReadGradientUnit(definition.Attribute("cy")?.Value, 0.5D);
            double r = ReadGradientUnit(definition.Attribute("r")?.Value, 0.5D);
            double fx = ReadGradientUnit(definition.Attribute("fx")?.Value, cx);
            double fy = ReadGradientUnit(definition.Attribute("fy")?.Value, cy);
            double fr = ReadGradientUnit(definition.Attribute("fr")?.Value, 0D);
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
                Dictionary<string, string> stopStyle = ReadDeclarations(stopElement.Attribute("style")?.Value);
                double offset = ReadGradientUnit(stopElement.Attribute("offset")?.Value, 0D);
                string? colorValue = stopElement.Attribute("stop-color")?.Value ?? (stopStyle.TryGetValue("stop-color", out string? styleColor) ? styleColor : null);
                double stopOpacity = ReadOpacity(stopElement.Attribute("stop-opacity")?.Value ?? (stopStyle.TryGetValue("stop-opacity", out string? styleOpacity) ? styleOpacity : null), 1D);
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

            if (value.StartsWith("rgb(", StringComparison.OrdinalIgnoreCase) && value.EndsWith(")", StringComparison.Ordinal)) {
                string inner = value.Substring(4, value.Length - 5);
                if (TryParseNumbers(inner, out List<double> parts) && parts.Count >= 3) {
                    color = ApplyAlpha(OfficeColor.FromRgb(ToByte(parts[0]), ToByte(parts[1]), ToByte(parts[2])), opacity);
                    return true;
                }
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

        private static bool IsUserSpaceGradient(XElement definition) =>
            string.Equals(definition.Attribute("gradientUnits")?.Value, "userSpaceOnUse", StringComparison.OrdinalIgnoreCase);

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

        private static double ClampUnit(double value) => value < 0D ? 0D : value > 1D ? 1D : value;
    }
}
