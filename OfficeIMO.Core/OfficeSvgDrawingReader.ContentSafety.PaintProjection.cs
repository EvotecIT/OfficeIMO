using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    private static bool HasKnownIncompleteSvgPaintProjection(XElement root) {
        XNamespace svgNamespace = root.Name.Namespace;
        if (HasAmbiguousSvgPaintServerDefinitions(root, svgNamespace)) return true;
        var references = new SvgElementReferenceRegistry(SvgDefinitionRegistry.Create(root));
        return root.DescendantsAndSelf().Any(element => {
            if (element.Attribute(XNamespace.Xml + "base") != null) return true;
            if (!IsNativeSvgElement(element, svgNamespace)) return false;
            string localName = element.Name.LocalName;
            if (IsCaseMismatchedSvgPaintDefinitionName(localName) ||
                HasEncodedSvgPaintServerReference(element, localName) ||
                HasUnsupportedSvgPaintServerMode(element, localName) ||
                HasUnsupportedSvgShapeGeometry(element, localName) ||
                HasUnisolatedSvgGroupOpacity(element, localName) ||
                HasUnsupportedSvgPaintStyle(element)) return true;
            if (HasActiveSvgPresentationProperty(element, "filter") ||
                HasActiveSvgPresentationProperty(element, "mask") ||
                HasActiveSvgPresentationProperty(element, "marker-start") ||
                HasActiveSvgPresentationProperty(element, "marker-mid") ||
                HasActiveSvgPresentationProperty(element, "marker-end")) return true;
            if (localName.Equals("textPath", StringComparison.Ordinal)) return true;
            if (localName.Equals("use", StringComparison.Ordinal)) {
                if (!TryOptionalUseLength(element, "x", out _) ||
                    !TryOptionalUseLength(element, "y", out _)) return true;
                if (references.TryEnter(element, out string referenceId, out XElement? target)) {
                    try {
                        if (target!.Name.LocalName.Equals("symbol", StringComparison.Ordinal) &&
                            (!TrySymbolLength(element, target, "width", 1D, out _) ||
                             !TrySymbolLength(element, target, "height", 1D, out _))) return true;
                    } finally {
                        references.Exit(referenceId);
                    }
                }
            }
            if (localName.Equals("foreignObject", StringComparison.Ordinal)) return true;
            if (localName.Equals("pattern", StringComparison.Ordinal)) {
                return element.Attribute("viewBox") != null ||
                    element.Attribute("preserveAspectRatio") != null;
            }
            if (!localName.Equals("image", StringComparison.Ordinal)) return false;
            XAttribute[] hrefs = element.Attributes()
                .Where(attribute => attribute.Name.LocalName.Equals("href", StringComparison.Ordinal))
                .ToArray();
            return hrefs.Length != 1 ||
                !TryDecodeEmbeddedRasterImage(hrefs[0].Value, out _, out _, out _) ||
                !TryViewportLength(element, "width", 1D, out double width) ||
                !TryViewportLength(element, "height", 1D, out double height) ||
                width <= 0D ||
                height <= 0D ||
                !TryParsePreserveAspectRatio(element.Attribute("preserveAspectRatio")?.Value, out _, out _);
        });
    }

    private static bool HasUnsupportedSvgShapeGeometry(XElement element, string localName) => localName switch {
        "rect" => HasUnsupportedSvgViewportLength(element, "x") ||
                  HasUnsupportedSvgViewportLength(element, "y") ||
                  HasUnsupportedSvgViewportLength(element, "width") ||
                  HasUnsupportedSvgViewportLength(element, "height") ||
                  HasUnsupportedSvgViewportLength(element, "rx") ||
                  HasUnsupportedSvgViewportLength(element, "ry"),
        "circle" => HasUnsupportedSvgViewportLength(element, "cx") ||
                    HasUnsupportedSvgViewportLength(element, "cy") ||
                    HasUnsupportedSvgViewportLength(element, "r"),
        "ellipse" => HasUnsupportedSvgViewportLength(element, "cx") ||
                     HasUnsupportedSvgViewportLength(element, "cy") ||
                     HasUnsupportedSvgViewportLength(element, "rx") ||
                     HasUnsupportedSvgViewportLength(element, "ry"),
        "line" => HasUnsupportedSvgViewportLength(element, "x1") ||
                  HasUnsupportedSvgViewportLength(element, "y1") ||
                  HasUnsupportedSvgViewportLength(element, "x2") ||
                  HasUnsupportedSvgViewportLength(element, "y2"),
        "polygon" or "polyline" => HasUnsupportedSvgPoints(element),
        "path" => HasUnsupportedSvgPathData(element),
        _ => false
    };

    private static bool HasUnsupportedSvgViewportLength(XElement element, string name) {
        string? value = element.Attribute(name)?.Value;
        return value != null &&
            (ContainsNonSvgCssWhitespace(value) || !TryViewportLength(value, 1D, out _, out _));
    }

    private static bool HasUnsupportedSvgPoints(XElement element) {
        string? points = element.Attribute("points")?.Value;
        return !string.IsNullOrWhiteSpace(points) &&
            !TryParseNumberList(points, MaximumSvgPathCommands * 2, out _);
    }

    private static bool HasUnsupportedSvgPathData(XElement element) {
        string? data = element.Attribute("d")?.Value;
        return !string.IsNullOrWhiteSpace(data) &&
            !OfficeSvgPathDataParser.TryParse(data, MaximumSvgPathCommands, out _, out _);
    }

    private static bool HasUnisolatedSvgGroupOpacity(XElement element, string localName) {
        if (!element.HasElements && localName is not ("text" or "tspan" or "use")) return false;
        string? value = ReadPresentationProperty(element, "opacity");
        if (string.IsNullOrWhiteSpace(value)) return false;
        return !TryUnit(value!, out double opacity) || opacity > 0D && opacity < 1D;
    }

    private static bool HasUnsupportedSvgPaintStyle(XElement element) {
        string? strokeWidth = ReadPresentationProperty(element, "stroke-width");
        if (!string.IsNullOrWhiteSpace(strokeWidth) &&
            (!TrySvgLength(strokeWidth, out double width) || width < 0D)) return true;
        if (HasUnsupportedSvgPresentationPaint(ReadPresentationProperty(element, "fill")) ||
            HasUnsupportedSvgPresentationPaint(ReadPresentationProperty(element, "stroke"))) return true;
        string? blend = ReadPresentationProperty(element, "mix-blend-mode");
        if (!string.IsNullOrWhiteSpace(blend) &&
            !TrimSvgCssWhitespace(blend!).Equals("normal", StringComparison.OrdinalIgnoreCase)) return true;
        string? vectorEffect = ReadPresentationProperty(element, "vector-effect");
        return !string.IsNullOrWhiteSpace(vectorEffect) &&
            !TrimSvgCssWhitespace(vectorEffect!).Equals("none", StringComparison.OrdinalIgnoreCase);
    }

    private static bool HasUnsupportedSvgPaintServerMode(XElement element, string localName) => localName switch {
        "linearGradient" or "radialGradient" =>
            HasNonExactSvgEnum(element, "gradientUnits", "objectBoundingBox", "userSpaceOnUse") ||
            HasNonExactSvgEnum(element, "spreadMethod", "pad", "reflect", "repeat"),
        "pattern" =>
            HasNonExactSvgEnum(element, "patternUnits", "objectBoundingBox", "userSpaceOnUse") ||
            HasNonExactSvgEnum(element, "patternContentUnits", "objectBoundingBox", "userSpaceOnUse"),
        _ => false
    };

    private static bool HasEncodedSvgPaintServerReference(XElement element, string localName) =>
        localName is "linearGradient" or "radialGradient" or "pattern" &&
        element.Attributes().Any(attribute =>
            attribute.Name.LocalName.Equals("href", StringComparison.Ordinal) &&
            attribute.Value.IndexOf('%') >= 0);

    private static bool HasNonExactSvgEnum(XElement element, string name, params string[] validValues) {
        string? value = element.Attribute(name)?.Value;
        return value != null && !validValues.Contains(value, StringComparer.Ordinal);
    }

    private static bool HasActiveSvgPresentationProperty(XElement element, string propertyName) {
        string? value = ReadPresentationProperty(element, propertyName);
        return !string.IsNullOrWhiteSpace(value) &&
            !TrimSvgCssWhitespace(value!).Equals("none", StringComparison.OrdinalIgnoreCase);
    }

    private static bool HasAmbiguousSvgPaintServerDefinitions(XElement root, XNamespace svgNamespace) {
        foreach (IGrouping<string, XElement> definitions in root.Descendants()
                     .Where(element => IsNativeSvgElement(element, svgNamespace))
                     .Select(element => new { Element = element, Id = ReadRasterElementId(element) })
                     .Where(item => item.Id != null)
                     .GroupBy(item => item.Id!, item => item.Element, StringComparer.Ordinal)) {
            if (definitions.Skip(1).Any() && definitions.Any(element =>
                    element.Name.LocalName is "linearGradient" or "radialGradient" or "pattern")) return true;
        }
        return false;
    }

    private static bool IsCaseMismatchedSvgPaintDefinitionName(string localName) =>
        (localName.Equals("linearGradient", StringComparison.OrdinalIgnoreCase) &&
         !localName.Equals("linearGradient", StringComparison.Ordinal)) ||
        (localName.Equals("radialGradient", StringComparison.OrdinalIgnoreCase) &&
         !localName.Equals("radialGradient", StringComparison.Ordinal)) ||
        (localName.Equals("pattern", StringComparison.OrdinalIgnoreCase) &&
         !localName.Equals("pattern", StringComparison.Ordinal)) ||
        (localName.Equals("stop", StringComparison.OrdinalIgnoreCase) &&
         !localName.Equals("stop", StringComparison.Ordinal));
}
