using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    private const int MaximumGradientReferenceDepth = 16;
    private const int MaximumGradientStops = 256;

    private static bool TryPaint(string value, SvgPaintServerRegistry paintServers, OfficeColor currentColor, out SvgResolvedPaint paint) {
        paint = default;
        if (value.Equals("none", StringComparison.OrdinalIgnoreCase)) return true;
        if (value.Equals("currentcolor", StringComparison.OrdinalIgnoreCase)) {
            paint = new SvgResolvedPaint(currentColor);
            return true;
        }
        if (value.StartsWith("url(", StringComparison.OrdinalIgnoreCase)) {
            return paintServers.TryResolve(value, out paint);
        }
        if (!TrySvgColor(value, out OfficeColor color)) return false;
        paint = new SvgResolvedPaint(color);
        return true;
    }

    private static bool TrySvgColor(string value, out OfficeColor color) {
        if (OfficeColor.TryParse(value, out color)) return true;

        color = default;
        string normalized = value.Trim();
        bool hasAlpha = normalized.StartsWith("rgba(", StringComparison.OrdinalIgnoreCase);
        bool hasRgb = normalized.StartsWith("rgb(", StringComparison.OrdinalIgnoreCase);
        if ((!hasAlpha && !hasRgb) || !normalized.EndsWith(")", StringComparison.Ordinal)) return false;

        int prefixLength = hasAlpha ? 5 : 4;
        string content = normalized.Substring(prefixLength, normalized.Length - prefixLength - 1);
        string[] components;
        string? alphaComponent = null;
        if (content.IndexOf(',') >= 0) {
            string[] commaComponents = content.Split(',');
            if (commaComponents.Length != (hasAlpha ? 4 : 3)) return false;
            components = new[] { commaComponents[0], commaComponents[1], commaComponents[2] };
            if (hasAlpha) alphaComponent = commaComponents[3];
        } else {
            string[] alphaParts = content.Split('/');
            if (alphaParts.Length > 2) return false;
            components = alphaParts[0].Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            if (alphaParts.Length == 2) alphaComponent = alphaParts[1];
        }

        if (components.Length != 3
            || !TrySvgColorChannel(components[0], out byte red)
            || !TrySvgColorChannel(components[1], out byte green)
            || !TrySvgColorChannel(components[2], out byte blue)) return false;

        byte alpha = 255;
        if (alphaComponent != null && !TrySvgAlphaChannel(alphaComponent, out alpha)) return false;
        color = OfficeColor.FromRgba(red, green, blue, alpha);
        return true;
    }

    private static bool TrySvgColorChannel(string text, out byte channel) {
        channel = 0;
        string normalized = text.Trim();
        bool percentage = normalized.EndsWith("%", StringComparison.Ordinal);
        if (percentage) normalized = normalized.Substring(0, normalized.Length - 1).Trim();
        if (!double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out double value)
            || double.IsNaN(value)
            || double.IsInfinity(value)
            || value < 0D
            || value > (percentage ? 100D : 255D)) return false;
        channel = (byte)Math.Round(percentage ? value * 255D / 100D : value);
        return true;
    }

    private static bool TrySvgAlphaChannel(string text, out byte alpha) {
        alpha = 0;
        string normalized = text.Trim();
        bool percentage = normalized.EndsWith("%", StringComparison.Ordinal);
        if (percentage) normalized = normalized.Substring(0, normalized.Length - 1).Trim();
        if (!double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out double value)
            || double.IsNaN(value)
            || double.IsInfinity(value)
            || value < 0D
            || value > (percentage ? 100D : 1D)) return false;
        alpha = (byte)Math.Round(percentage ? value * 255D / 100D : value * 255D);
        return true;
    }

    private readonly struct SvgResolvedPaint {
        internal OfficeColor? Color { get; }
        internal OfficeLinearGradient? LinearGradient { get; }
        internal OfficeRadialGradient? RadialGradient { get; }
        internal SvgGradientDefinition? DeferredGradient { get; }

        internal SvgResolvedPaint(OfficeColor color) {
            Color = color;
            LinearGradient = null;
            RadialGradient = null;
            DeferredGradient = null;
        }

        internal SvgResolvedPaint(OfficeLinearGradient gradient) {
            Color = null;
            LinearGradient = gradient;
            RadialGradient = null;
            DeferredGradient = null;
        }

        internal SvgResolvedPaint(OfficeRadialGradient gradient) {
            Color = null;
            LinearGradient = null;
            RadialGradient = gradient;
            DeferredGradient = null;
        }

        internal SvgResolvedPaint(SvgGradientDefinition gradient) {
            Color = null;
            LinearGradient = null;
            RadialGradient = null;
            DeferredGradient = gradient;
        }
    }

    private sealed class SvgPaintServerRegistry {
        private readonly SvgDefinitionRegistry _definitions;

        internal SvgPaintServerRegistry(SvgDefinitionRegistry definitions) {
            _definitions = definitions;
        }

        internal bool TryResolve(string value, out SvgResolvedPaint paint) {
            paint = default;
            if (!TryReadLocalReference(value, requireUrl: true, out string id)
                || !_definitions.TryGetUnique(id, out XElement? element)
                || (!element!.Name.LocalName.Equals("linearGradient", StringComparison.OrdinalIgnoreCase)
                    && !element.Name.LocalName.Equals("radialGradient", StringComparison.OrdinalIgnoreCase))) return false;

            var resolving = new HashSet<string>(StringComparer.Ordinal);
            if (!TryResolveDefinition(id, element, resolving, 0, out SvgGradientDefinition? definition)) return false;
            try {
                if (definition!.UserSpaceOnUse || definition.GradientTransform != OfficeTransform.Identity || definition.SpreadMode != SvgGradientSpreadMode.Pad) {
                    paint = new SvgResolvedPaint(definition);
                } else if (definition.Kind == SvgGradientKind.Linear) {
                    paint = new SvgResolvedPaint(new OfficeLinearGradient(
                        definition.X1.Value,
                        definition.Y1.Value,
                        definition.X2.Value,
                        definition.Y2.Value,
                        definition.Stops));
                } else {
                    paint = new SvgResolvedPaint(new OfficeRadialGradient(
                        definition.X1.Value,
                        definition.Y1.Value,
                        definition.Radius1.Value,
                        definition.X2.Value,
                        definition.Y2.Value,
                        definition.Radius2.Value,
                        definition.Stops));
                }
                return true;
            } catch (ArgumentException) {
                return false;
            }
        }

        private bool TryResolveDefinition(
            string id,
            XElement element,
            ISet<string> resolving,
            int depth,
            out SvgGradientDefinition? definition) {
            definition = null;
            if (depth >= MaximumGradientReferenceDepth || !resolving.Add(id)) return false;
            try {
                SvgGradientKind kind = element.Name.LocalName.Equals("linearGradient", StringComparison.OrdinalIgnoreCase)
                    ? SvgGradientKind.Linear
                    : SvgGradientKind.Radial;
                SvgGradientDefinition? inherited = null;
                XAttribute? href = element.Attributes().FirstOrDefault(attribute => attribute.Name.LocalName.Equals("href", StringComparison.OrdinalIgnoreCase));
                if (href != null) {
                    if (!TryReadLocalReference(href.Value, requireUrl: false, out string inheritedId)
                        || !_definitions.TryGetUnique(inheritedId, out XElement? inheritedElement)
                        || inheritedElement == null
                        || !inheritedElement.Name.LocalName.Equals(element.Name.LocalName, StringComparison.OrdinalIgnoreCase)
                        || !TryResolveDefinition(inheritedId, inheritedElement, resolving, depth + 1, out inherited)) return false;
                }

                if (!UsesSupportedGradientOptions(element)) return false;
                if (!TryResolveGradientUnits(element, inherited, out bool userSpaceOnUse)) return false;
                if (!TryResolveGradientTransform(element, inherited, out OfficeTransform gradientTransform)) return false;
                if (!TryResolveSpreadMode(element, inherited, out SvgGradientSpreadMode spreadMode)) return false;
                IReadOnlyList<OfficeGradientStop>? stops = null;
                if (element.Elements().Any(child => child.Name.LocalName.Equals("stop", StringComparison.OrdinalIgnoreCase))) {
                    if (!TryReadStops(element, out stops)) return false;
                } else if (inherited != null) {
                    stops = inherited.Stops;
                }
                if (stops == null) return false;

                if (kind == SvgGradientKind.Linear) {
                    SvgGradientCoordinate defaultX1 = inherited?.X1 ?? SvgGradientCoordinate.CreateDefault(0D);
                    SvgGradientCoordinate defaultY1 = inherited?.Y1 ?? SvgGradientCoordinate.CreateDefault(0D);
                    SvgGradientCoordinate defaultX2 = inherited?.X2 ?? SvgGradientCoordinate.CreateDefault(1D);
                    SvgGradientCoordinate defaultY2 = inherited?.Y2 ?? SvgGradientCoordinate.CreateDefault(0D);
                    if (!TryCoordinate(element, "x1", defaultX1, allowOutsideUnit: false, userSpaceOnUse, out SvgGradientCoordinate x1)
                        || !TryCoordinate(element, "y1", defaultY1, allowOutsideUnit: false, userSpaceOnUse, out SvgGradientCoordinate y1)
                        || !TryCoordinate(element, "x2", defaultX2, allowOutsideUnit: false, userSpaceOnUse, out SvgGradientCoordinate x2)
                        || !TryCoordinate(element, "y2", defaultY2, allowOutsideUnit: false, userSpaceOnUse, out SvgGradientCoordinate y2)
                        || (x1.Equals(x2) && y1.Equals(y2))) return false;
                    definition = SvgGradientDefinition.Linear(x1, y1, x2, y2, stops, userSpaceOnUse, gradientTransform, spreadMode);
                    return true;
                }

                SvgGradientCoordinate defaultCenterX = inherited?.X2 ?? SvgGradientCoordinate.CreateDefault(0.5D);
                SvgGradientCoordinate defaultCenterY = inherited?.Y2 ?? SvgGradientCoordinate.CreateDefault(0.5D);
                SvgGradientCoordinate defaultRadius = inherited?.Radius2 ?? SvgGradientCoordinate.CreateDefault(0.5D);
                if (!TryCoordinate(element, "cx", defaultCenterX, allowOutsideUnit: true, userSpaceOnUse, out SvgGradientCoordinate centerX)
                    || !TryCoordinate(element, "cy", defaultCenterY, allowOutsideUnit: true, userSpaceOnUse, out SvgGradientCoordinate centerY)
                    || !TryCoordinate(element, "r", defaultRadius, allowOutsideUnit: true, userSpaceOnUse, out SvgGradientCoordinate radius)
                    || !TryCoordinate(element, "fx", inherited?.X1 ?? centerX, allowOutsideUnit: true, userSpaceOnUse, out SvgGradientCoordinate focalX)
                    || !TryCoordinate(element, "fy", inherited?.Y1 ?? centerY, allowOutsideUnit: true, userSpaceOnUse, out SvgGradientCoordinate focalY)
                    || !TryCoordinate(element, "fr", inherited?.Radius1 ?? SvgGradientCoordinate.CreateDefault(0D), allowOutsideUnit: true, userSpaceOnUse, out SvgGradientCoordinate focalRadius)
                    || radius.Value <= 0D
                    || focalRadius.Value < 0D
                    || (focalX.Equals(centerX) && focalY.Equals(centerY) && focalRadius.Equals(radius))) return false;
                if (!userSpaceOnUse && focalRadius.Value > radius.Value) return false;
                definition = SvgGradientDefinition.Radial(focalX, focalY, focalRadius, centerX, centerY, radius, stops, userSpaceOnUse, gradientTransform, spreadMode);
                return true;
            } finally {
                resolving.Remove(id);
            }
        }

        private static bool UsesSupportedGradientOptions(XElement element) {
            return true;
        }

        private static bool TryResolveSpreadMode(
            XElement element,
            SvgGradientDefinition? inherited,
            out SvgGradientSpreadMode spreadMode) {
            string? value = element.Attribute("spreadMethod")?.Value.Trim();
            if (string.IsNullOrEmpty(value)) {
                spreadMode = inherited?.SpreadMode ?? SvgGradientSpreadMode.Pad;
                return true;
            }
            if (value!.Equals("pad", StringComparison.OrdinalIgnoreCase)) spreadMode = SvgGradientSpreadMode.Pad;
            else if (value.Equals("repeat", StringComparison.OrdinalIgnoreCase)) spreadMode = SvgGradientSpreadMode.Repeat;
            else if (value.Equals("reflect", StringComparison.OrdinalIgnoreCase)) spreadMode = SvgGradientSpreadMode.Reflect;
            else {
                spreadMode = default;
                return false;
            }
            return true;
        }

        private static bool TryResolveGradientTransform(
            XElement element,
            SvgGradientDefinition? inherited,
            out OfficeTransform transform) {
            string? value = element.Attribute("gradientTransform")?.Value;
            if (string.IsNullOrWhiteSpace(value)) {
                transform = inherited?.GradientTransform ?? OfficeTransform.Identity;
                return true;
            }
            return OfficeSvgTransformParser.TryParse(value, out transform);
        }

        private static bool TryResolveGradientUnits(XElement element, SvgGradientDefinition? inherited, out bool userSpaceOnUse) {
            string? units = element.Attribute("gradientUnits")?.Value.Trim();
            if (string.IsNullOrEmpty(units)) {
                userSpaceOnUse = inherited?.UserSpaceOnUse ?? false;
                return true;
            }
            if (units!.Equals("objectBoundingBox", StringComparison.OrdinalIgnoreCase)) {
                userSpaceOnUse = false;
                return true;
            }
            if (units.Equals("userSpaceOnUse", StringComparison.OrdinalIgnoreCase)) {
                userSpaceOnUse = true;
                return true;
            }
            userSpaceOnUse = false;
            return false;
        }

        private static bool TryReadStops(XElement gradient, out IReadOnlyList<OfficeGradientStop>? stops) {
            stops = null;
            XElement[] elements = gradient.Elements()
                .Where(element => element.Name.LocalName.Equals("stop", StringComparison.OrdinalIgnoreCase))
                .ToArray();
            if (elements.Length == 0 || elements.Length > MaximumGradientStops) return false;
            if (!TryResolveCurrentColor(gradient, out OfficeColor inheritedCurrentColor)) return false;

            var parsed = new List<OfficeGradientStop>(elements.Length + 2);
            double previous = -1D;
            foreach (XElement element in elements) {
                if (!TryStopOffset(element.Attribute("offset")?.Value, out double offset)) return false;
                offset = Math.Max(previous, offset);
                if (!TryStopColor(element, inheritedCurrentColor, out OfficeColor color)) return false;
                parsed.Add(new OfficeGradientStop(offset, color));
                previous = offset;
            }

            if (parsed.Count == 1) {
                OfficeColor color = parsed[0].Color;
                stops = new[] { new OfficeGradientStop(0D, color), new OfficeGradientStop(1D, color) };
                return true;
            }
            if (parsed[0].Offset > 0D) parsed.Insert(0, new OfficeGradientStop(0D, parsed[0].Color));
            if (parsed[parsed.Count - 1].Offset < 1D) parsed.Add(new OfficeGradientStop(1D, parsed[parsed.Count - 1].Color));
            stops = parsed;
            return true;
        }

        private static bool TryStopOffset(string? value, out double offset) {
            if (string.IsNullOrWhiteSpace(value)) {
                offset = 0D;
                return true;
            }
            return TryUnitOrPercentage(value!, clamp: true, out offset);
        }

        private static bool TryStopColor(XElement element, OfficeColor inheritedCurrentColor,
            out OfficeColor color) {
            color = OfficeColor.Black;
            OfficeColor currentColor = inheritedCurrentColor;
            string? currentColorText = element.Attribute("color")?.Value;
            string? colorText = element.Attribute("stop-color")?.Value;
            string? opacityText = element.Attribute("stop-opacity")?.Value;
            string? declarations = element.Attribute("style")?.Value;
            if (!string.IsNullOrWhiteSpace(declarations)) {
                foreach (string declaration in declarations!.Split(';')) {
                    int colon = declaration.IndexOf(':');
                    if (colon <= 0) continue;
                    string name = declaration.Substring(0, colon).Trim();
                    string value = declaration.Substring(colon + 1).Trim();
                    if (name.Equals("stop-color", StringComparison.OrdinalIgnoreCase)) colorText = value;
                    else if (name.Equals("stop-opacity", StringComparison.OrdinalIgnoreCase)) opacityText = value;
                    else if (name.Equals("color", StringComparison.OrdinalIgnoreCase)) currentColorText = value;
                }
            }

            if (!string.IsNullOrWhiteSpace(currentColorText) &&
                !currentColorText!.Trim().Equals("currentcolor", StringComparison.OrdinalIgnoreCase) &&
                !TrySvgColor(currentColorText.Trim(), out currentColor)) return false;

            if (string.IsNullOrWhiteSpace(colorText)) color = OfficeColor.Black;
            else if (colorText!.Trim().Equals("currentcolor", StringComparison.OrdinalIgnoreCase)) color = currentColor;
            else if (!TrySvgColor(colorText.Trim(), out color)) return false;

            double opacity = 1D;
            if (!string.IsNullOrWhiteSpace(opacityText) && !TryUnitOrPercentage(opacityText!, clamp: true, out opacity)) return false;
            color = OfficeColor.FromRgba(color.R, color.G, color.B, (byte)Math.Round(color.A * opacity));
            return true;
        }

        private static bool TryResolveCurrentColor(XElement element, out OfficeColor color) {
            color = OfficeColor.Black;
            foreach (XElement candidate in element.AncestorsAndSelf().Reverse()) {
                string? value = candidate.Attribute("color")?.Value;
                string? declarations = candidate.Attribute("style")?.Value;
                if (!string.IsNullOrWhiteSpace(declarations)) {
                    foreach (string declaration in declarations!.Split(';')) {
                        int colon = declaration.IndexOf(':');
                        if (colon <= 0 || !declaration.Substring(0, colon).Trim().Equals("color", StringComparison.OrdinalIgnoreCase)) continue;
                        value = declaration.Substring(colon + 1).Trim();
                    }
                }

                if (string.IsNullOrWhiteSpace(value) || value!.Trim().Equals("currentcolor", StringComparison.OrdinalIgnoreCase)) continue;
                if (!TrySvgColor(value.Trim(), out color)) return false;
            }
            return true;
        }

        private static bool TryCoordinate(
            XElement element,
            string name,
            SvgGradientCoordinate fallback,
            bool allowOutsideUnit,
            bool userSpaceOnUse,
            out SvgGradientCoordinate coordinate) {
            string? text = element.Attribute(name)?.Value;
            if (string.IsNullOrWhiteSpace(text)) {
                coordinate = fallback;
                return true;
            }
            string normalized = text!.Trim();
            bool percentage = normalized.EndsWith("%", StringComparison.Ordinal);
            if (percentage) normalized = normalized.Substring(0, normalized.Length - 1).Trim();
            else if (normalized.EndsWith("px", StringComparison.OrdinalIgnoreCase)) normalized = normalized.Substring(0, normalized.Length - 2).Trim();
            if (!double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out double value)
                || double.IsNaN(value)
                || double.IsInfinity(value)) {
                coordinate = default;
                return false;
            }
            if (percentage) value /= 100D;
            coordinate = new SvgGradientCoordinate(value, percentage);
            return userSpaceOnUse || allowOutsideUnit || (value >= 0D && value <= 1D);
        }

        private static bool TryUnitOrPercentage(string text, bool clamp, out double value) {
            string normalized = text.Trim();
            bool percentage = normalized.EndsWith("%", StringComparison.Ordinal);
            if (percentage) normalized = normalized.Substring(0, normalized.Length - 1).Trim();
            if (!double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out value)
                || double.IsNaN(value)
                || double.IsInfinity(value)) return false;
            if (percentage) value /= 100D;
            if (clamp) value = value < 0D ? 0D : value > 1D ? 1D : value;
            return true;
        }

        private static bool TryReadLocalReference(string text, bool requireUrl, out string id) {
            id = string.Empty;
            string normalized = text.Trim();
            if (requireUrl) {
                if (!normalized.StartsWith("url(", StringComparison.OrdinalIgnoreCase) || !normalized.EndsWith(")", StringComparison.Ordinal)) return false;
                normalized = normalized.Substring(4, normalized.Length - 5).Trim().Trim('\'', '"');
            }
            if (normalized.Length < 2 || normalized[0] != '#') return false;
            id = normalized.Substring(1);
            return id.Length > 0 && id.IndexOfAny(new[] { ' ', '\t', '\r', '\n', '#', '(', ')' }) < 0;
        }
    }

    private enum SvgGradientKind {
        Linear,
        Radial
    }

    private enum SvgGradientSpreadMode {
        Pad,
        Repeat,
        Reflect
    }

    private sealed partial class SvgGradientDefinition {
        internal SvgGradientKind Kind { get; private set; }
        internal SvgGradientCoordinate X1 { get; private set; }
        internal SvgGradientCoordinate Y1 { get; private set; }
        internal SvgGradientCoordinate Radius1 { get; private set; }
        internal SvgGradientCoordinate X2 { get; private set; }
        internal SvgGradientCoordinate Y2 { get; private set; }
        internal SvgGradientCoordinate Radius2 { get; private set; }
        internal bool UserSpaceOnUse { get; private set; }
        internal OfficeTransform GradientTransform { get; private set; }
        internal SvgGradientSpreadMode SpreadMode { get; private set; }
        internal IReadOnlyList<OfficeGradientStop> Stops { get; private set; } = Array.Empty<OfficeGradientStop>();

        internal static SvgGradientDefinition Linear(
            SvgGradientCoordinate x1,
            SvgGradientCoordinate y1,
            SvgGradientCoordinate x2,
            SvgGradientCoordinate y2,
            IReadOnlyList<OfficeGradientStop> stops,
            bool userSpaceOnUse,
            OfficeTransform gradientTransform,
            SvgGradientSpreadMode spreadMode) =>
            new SvgGradientDefinition { Kind = SvgGradientKind.Linear, X1 = x1, Y1 = y1, X2 = x2, Y2 = y2, Stops = stops, UserSpaceOnUse = userSpaceOnUse, GradientTransform = gradientTransform, SpreadMode = spreadMode };

        internal static SvgGradientDefinition Radial(
            SvgGradientCoordinate x1,
            SvgGradientCoordinate y1,
            SvgGradientCoordinate radius1,
            SvgGradientCoordinate x2,
            SvgGradientCoordinate y2,
            SvgGradientCoordinate radius2,
            IReadOnlyList<OfficeGradientStop> stops,
            bool userSpaceOnUse,
            OfficeTransform gradientTransform,
            SvgGradientSpreadMode spreadMode) =>
            new SvgGradientDefinition { Kind = SvgGradientKind.Radial, X1 = x1, Y1 = y1, Radius1 = radius1, X2 = x2, Y2 = y2, Radius2 = radius2, Stops = stops, UserSpaceOnUse = userSpaceOnUse, GradientTransform = gradientTransform, SpreadMode = spreadMode };

        internal bool TryCreateForShape(
            OfficeShape shape,
            double shapeX,
            double shapeY,
            double viewportWidth,
            double viewportHeight,
            double viewX,
            double viewY,
            out OfficeLinearGradient? linear,
            out OfficeRadialGradient? radial) {
            linear = null;
            radial = null;
            if (shape.Width <= 0D || shape.Height <= 0D || viewportWidth <= 0D || viewportHeight <= 0D) return false;
            try {
                OfficePoint first = ResolvePoint(X1, Y1, viewportWidth, viewportHeight, viewX, viewY);
                OfficePoint second = ResolvePoint(X2, Y2, viewportWidth, viewportHeight, viewX, viewY);
                first = GradientTransform.TransformPoint(first);
                second = GradientTransform.TransformPoint(second);
                double x1 = NormalizeAxis(first.X, shapeX, shape.Width, viewX);
                double y1 = NormalizeAxis(first.Y, shapeY, shape.Height, viewY);
                double x2 = NormalizeAxis(second.X, shapeX, shape.Width, viewX);
                double y2 = NormalizeAxis(second.Y, shapeY, shape.Height, viewY);
                if (Kind == SvgGradientKind.Linear) {
                    return TryCreateLinearSpread(x1, y1, x2, y2, out linear);
                }

                if (SpreadMode != SvgGradientSpreadMode.Pad) return false;
                if (Math.Abs(GradientTransform.M12) > 0.0000001D || Math.Abs(GradientTransform.M21) > 0.0000001D) return false;
                double diagonal = Math.Sqrt((viewportWidth * viewportWidth) + (viewportHeight * viewportHeight)) / Math.Sqrt(2D);
                double radius1 = ResolveRadius(Radius1, diagonal, UserSpaceOnUse);
                double radius2 = ResolveRadius(Radius2, diagonal, UserSpaceOnUse);
                double radius1X = radius1 * Math.Abs(GradientTransform.M11);
                double radius1Y = radius1 * Math.Abs(GradientTransform.M22);
                double radius2X = radius2 * Math.Abs(GradientTransform.M11);
                double radius2Y = radius2 * Math.Abs(GradientTransform.M22);
                if (radius2X <= 0D || radius2Y <= 0D || radius1X < 0D || radius1Y < 0D || radius1X > radius2X || radius1Y > radius2Y) return false;
                radial = new OfficeRadialGradient(
                    x1,
                    y1,
                    NormalizeRadius(radius1X, shape.Width, UserSpaceOnUse),
                    NormalizeRadius(radius1Y, shape.Height, UserSpaceOnUse),
                    x2,
                    y2,
                    NormalizeRadius(radius2X, shape.Width, UserSpaceOnUse),
                    NormalizeRadius(radius2Y, shape.Height, UserSpaceOnUse),
                    Stops);
                return true;
            } catch (ArgumentException) {
                return false;
            }
        }

        private OfficePoint ResolvePoint(
            SvgGradientCoordinate x,
            SvgGradientCoordinate y,
            double viewportWidth,
            double viewportHeight,
            double viewX,
            double viewY) {
            if (!UserSpaceOnUse) return new OfficePoint(x.Value, y.Value);
            return new OfficePoint(
                x.IsPercentage ? viewX + (x.Value * viewportWidth) : x.Value,
                y.IsPercentage ? viewY + (y.Value * viewportHeight) : y.Value);
        }

        private double NormalizeAxis(double value, double shapeOrigin, double shapeSize, double viewOrigin) =>
            UserSpaceOnUse ? ((value - viewOrigin) - shapeOrigin) / shapeSize : value;

        private static double ResolveRadius(SvgGradientCoordinate coordinate, double viewportDiagonal, bool userSpaceOnUse) =>
            userSpaceOnUse && coordinate.IsPercentage ? coordinate.Value * viewportDiagonal : coordinate.Value;

        private static double NormalizeRadius(double radius, double shapeSize, bool userSpaceOnUse) =>
            userSpaceOnUse ? radius / shapeSize : radius;

    }

    private readonly struct SvgGradientCoordinate : IEquatable<SvgGradientCoordinate> {
        internal double Value { get; }
        internal bool IsPercentage { get; }

        internal SvgGradientCoordinate(double value, bool isPercentage) {
            Value = value;
            IsPercentage = isPercentage;
        }

        internal static SvgGradientCoordinate CreateDefault(double value) => new SvgGradientCoordinate(value, isPercentage: true);

        public bool Equals(SvgGradientCoordinate other) => Value.Equals(other.Value) && IsPercentage == other.IsPercentage;
        public override bool Equals(object? obj) => obj is SvgGradientCoordinate other && Equals(other);
        public override int GetHashCode() => unchecked((Value.GetHashCode() * 397) ^ IsPercentage.GetHashCode());
    }
}
