using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    private static bool HasKnownIncompleteSvgPaintProjection(XElement sourceRoot, XElement root) {
        XNamespace svgNamespace = root.Name.Namespace;
        if (HasSvgUseShadowStyleProjectionRisk(sourceRoot)) return true;
        if (HasAmbiguousSvgPaintServerDefinitions(root, svgNamespace)) return true;
        SvgDefinitionRegistry definitions = SvgDefinitionRegistry.Create(root);
        var paintServers = new SvgPaintServerRegistry(definitions);
        ISet<XElement> activePaintElements = CollectSvgActivePaintElements(root, definitions);
        return root.DescendantsAndSelf().Any(element => {
            if (element.Attribute(XNamespace.Xml + "base") != null) return true;
            if (!IsNativeSvgElement(element, svgNamespace)) return false;
            string localName = element.Name.LocalName;
            // Native viewport and tile rendering can clip paint which a browser may expose
            // outside a nested SVG, referenced symbol, or pattern tile.
            if (localName is "symbol" or "pattern" ||
                localName.Equals("svg", StringComparison.Ordinal) && element.Parent != null) {
                string? overflow = ReadPresentationProperty(element, "overflow");
                if (!string.IsNullOrWhiteSpace(overflow) &&
                    !TrimSvgCssWhitespace(overflow!).Equals("hidden", StringComparison.OrdinalIgnoreCase)) return true;
            }
            if (IsCaseMismatchedSvgPaintDefinitionName(localName) ||
                HasEncodedSvgPaintServerReference(element, localName) ||
                HasUnsupportedSvgPaintServerMode(element, localName, activePaintElements.Contains(element)) ||
                HasUnsupportedSvgShapeGeometry(element, localName) ||
                HasUnisolatedSvgGroupOpacity(element, localName) ||
                activePaintElements.Contains(element) && HasUnsupportedSvgPaintStyle(element, paintServers)) return true;
            if (HasActiveSvgPresentationProperty(element, "filter") ||
                HasActiveSvgPresentationProperty(element, "mask") ||
                HasActiveSvgPresentationProperty(element, "marker-start") ||
                HasActiveSvgPresentationProperty(element, "marker-mid") ||
                HasActiveSvgPresentationProperty(element, "marker-end")) return true;
            if (localName.Equals("textPath", StringComparison.Ordinal)) return true;
            if (localName.Equals("use", StringComparison.Ordinal)) {
                // A local reference can recurse through another use beyond the native depth
                // limit, or acquire inherited paint after cloning into its shadow tree.
                return activePaintElements.Contains(element);
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

    private static ISet<XElement> CollectSvgActivePaintElements(XElement root, SvgDefinitionRegistry definitions) {
        var active = new HashSet<XElement>();
        var pending = new Stack<XElement>();
        pending.Push(root);
        while (pending.Count > 0) {
            XElement element = pending.Pop();
            if (!active.Add(element)) continue;
            foreach (string property in new[] { "fill", "stroke" }) {
                string? value = ReadPresentationProperty(element, property);
                if (value != null && TryReadBoundedSvgLocalUrlReference(value, out string reference)) {
                    EnqueueSvgPaintReference(reference, definitions, active, pending);
                }
            }
            foreach (XAttribute href in element.Attributes().Where(attribute =>
                         attribute.Name.LocalName.Equals("href", StringComparison.Ordinal))) {
                EnqueueSvgPaintReference(href.Value, definitions, active, pending);
            }
            foreach (XElement child in element.Elements()) {
                if (child.Name.LocalName is not ("defs" or "symbol" or "pattern" or "linearGradient" or "radialGradient")) {
                    pending.Push(child);
                }
            }
        }
        return active;
    }

    private static void EnqueueSvgPaintReference(
        string reference,
        SvgDefinitionRegistry definitions,
        ISet<XElement> active,
        Stack<XElement> pending) {
        if (reference.Length < 2 || reference[0] != '#' || reference.IndexOf('%') >= 0 ||
            !definitions.TryGetUnique(reference.Substring(1), out XElement? target)) return;
        for (XElement? ancestor = target!.Parent; ancestor != null; ancestor = ancestor.Parent) {
            active.Add(ancestor);
        }
        pending.Push(target);
    }

    private static bool HasSvgUseShadowStyleProjectionRisk(XElement root) {
        XNamespace svgNamespace = root.Name.Namespace;
        if (!root.DescendantsAndSelf().Any(element =>
                IsNativeSvgElement(element, svgNamespace) &&
                element.Name.LocalName.Equals("use", StringComparison.Ordinal))) return false;

        // The computed source tree is not the use shadow tree: inherited values, custom
        // properties, and selector matches can change when the referenced subtree is cloned.
        return root.DescendantsAndSelf().Any(element =>
            IsNativeSvgElement(element, svgNamespace) &&
            (element.Name.LocalName.Equals("style", StringComparison.Ordinal) && element.Value.Length > 0 ||
             element.Attributes().Any(attribute =>
                 attribute.Name.NamespaceName.Length == 0 &&
                 (attribute.Name.LocalName.Equals("style", StringComparison.Ordinal) &&
                      (attribute.Value.IndexOf("inherit", StringComparison.OrdinalIgnoreCase) >= 0 ||
                       attribute.Value.IndexOf("var(", StringComparison.OrdinalIgnoreCase) >= 0) ||
                  attribute.Value.Trim().Equals("inherit", StringComparison.OrdinalIgnoreCase) ||
                  IsSvgPresentationPropertyName(attribute.Name.LocalName) &&
                      attribute.Value.IndexOf("var(", StringComparison.OrdinalIgnoreCase) >= 0))));
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

    private static bool HasUnsupportedSvgPaintStyle(XElement element, SvgPaintServerRegistry paintServers) {
        string? strokeWidth = ReadPresentationProperty(element, "stroke-width");
        if (!string.IsNullOrWhiteSpace(strokeWidth) &&
            (!TrySvgLength(strokeWidth, out double width) || width < 0D)) return true;
        string? color = ReadPresentationProperty(element, "color");
        if (!string.IsNullOrWhiteSpace(color) &&
            !TrimSvgCssWhitespace(color!).Equals("currentColor", StringComparison.OrdinalIgnoreCase) &&
            !OfficeColor.TryParseCss(TrimSvgCssWhitespace(color!), out _)) return true;
        if (HasUnsupportedSvgPresentationPaint(ReadPresentationProperty(element, "fill"), requireCompleteNativeProjection: true) ||
            HasUnsupportedSvgPresentationPaint(ReadPresentationProperty(element, "stroke"), requireCompleteNativeProjection: true)) return true;
        foreach (string name in new[] {
            "opacity", "fill-opacity", "stroke-opacity", "stroke-dasharray", "stroke-dashoffset",
            "stroke-linecap", "stroke-linejoin", "stroke-miterlimit", "fill-rule"
        }) {
            string? value = ReadPresentationProperty(element, name);
            if (string.IsNullOrWhiteSpace(value)) continue;
            SvgPaintContext validation = SvgPaintContext.Default;
            int unsupported = 0;
            ApplyProperty(name, value, paintServers, ref validation, ref unsupported);
            if (unsupported != 0) return true;
        }
        string? blend = ReadPresentationProperty(element, "mix-blend-mode");
        if (!string.IsNullOrWhiteSpace(blend) &&
            !TrimSvgCssWhitespace(blend!).Equals("normal", StringComparison.OrdinalIgnoreCase)) return true;
        string? vectorEffect = ReadPresentationProperty(element, "vector-effect");
        return !string.IsNullOrWhiteSpace(vectorEffect) &&
            !TrimSvgCssWhitespace(vectorEffect!).Equals("none", StringComparison.OrdinalIgnoreCase);
    }

    private static bool HasUnsupportedSvgPaintServerMode(XElement element, string localName, bool active) => localName switch {
        "linearGradient" or "radialGradient" =>
            HasNonExactSvgEnum(element, "gradientUnits", "objectBoundingBox", "userSpaceOnUse") ||
            HasNonExactSvgEnum(element, "spreadMethod", "pad", "reflect", "repeat"),
        "pattern" =>
            HasNonExactSvgEnum(element, "patternUnits", "objectBoundingBox", "userSpaceOnUse") ||
            HasNonExactSvgEnum(element, "patternContentUnits", "objectBoundingBox", "userSpaceOnUse") ||
            active && (element.Attributes().Any(attribute => attribute.Name.LocalName.Equals("href", StringComparison.Ordinal)) ||
                HasUnsupportedSvgPatternGeometry(element)),
        _ => false
    };

    private static bool HasUnsupportedSvgPatternGeometry(XElement pattern) {
        bool userSpace = string.Equals(pattern.Attribute("patternUnits")?.Value, "userSpaceOnUse", StringComparison.Ordinal);
        foreach (string name in new[] { "x", "y", "width", "height" }) {
            string? value = pattern.Attribute(name)?.Value;
            if (value == null) continue;
            if (ContainsNonSvgCssWhitespace(value) ||
                (userSpace
                    ? !TryViewportLength(value, 1D, out _, out _)
                    : !TryPatternBoxFraction(value, 0D, out _))) return true;
        }
        return false;
    }

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
