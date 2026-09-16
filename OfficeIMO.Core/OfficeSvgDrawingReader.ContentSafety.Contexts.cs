using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.ContentSafety;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    private static bool TryResolveSvgContentSafetyNestedViewport(
        XElement element,
        OfficeTransform elementTransform,
        double parentViewX,
        double parentViewY,
        double parentViewWidth,
        double parentViewHeight,
        SvgContentSafetyDocument document,
        out OfficeTransform contentTransform,
        out double childViewX,
        out double childViewY,
        out double childViewWidth,
        out double childViewHeight) {
        contentTransform = elementTransform;
        childViewX = 0D;
        childViewY = 0D;
        childViewWidth = parentViewWidth;
        childViewHeight = parentViewHeight;
        double x = ReadViewportCoordinate(element, "x", parentViewX, parentViewWidth);
        double y = ReadViewportCoordinate(element, "y", parentViewY, parentViewHeight);
        if (!TryNestedViewportLength(element.Attribute("width")?.Value, parentViewWidth, out double width)
            || !TryNestedViewportLength(element.Attribute("height")?.Value, parentViewHeight, out double height)
            || !IsSupportedSvgViewport(width, height, document.MaximumViewportDimension, document.MaximumViewportPixels)) {
            return false;
        }

        string? viewBoxText = element.Attribute("viewBox")?.Value;
        if (string.IsNullOrWhiteSpace(viewBoxText)) {
            childViewWidth = width;
            childViewHeight = height;
        } else {
            if (!TryParseNumberList(viewBoxText, out IReadOnlyList<double> viewBox)
                || viewBox.Count != 4
                || !IsSupportedSvgViewport(viewBox[2], viewBox[3], document.MaximumViewportDimension, document.MaximumViewportPixels)) {
                return false;
            }
            childViewX = viewBox[0];
            childViewY = viewBox[1];
            childViewWidth = viewBox[2];
            childViewHeight = viewBox[3];
        }
        if (!TryParsePreserveAspectRatio(element.Attribute("preserveAspectRatio")?.Value,
                out SvgAspectAlignment alignment, out bool slice)) return false;
        contentTransform = ResolveViewportTransform(childViewWidth, childViewHeight, width, height, alignment, slice)
            .Then(OfficeTransform.Translate(x, y))
            .Then(elementTransform);
        return IsSupportedSvgTransform(contentTransform);
    }

    private static bool TryClassifySvgNonPrimary(
        SvgContentSafetyCandidate candidate,
        out SvgContentSafetyConcealment concealment,
        out OfficeContentCleanupCapability cleanupCapability) {
        XElement? owner = candidate.ComputedElement.AncestorsAndSelf().FirstOrDefault(element => {
            string name = element.Name.LocalName;
            return name is "title" or "desc" or "script" or "style" or "metadata";
        });
        if (!candidate.IsNativeSvg) {
            concealment = new SvgContentSafetyConcealment(
                OfficeContentConcealmentKind.NonPrimaryContent,
                "Foreign-namespace XML text is machine-readable extension content and is report-only because it is not native SVG paint.",
                OfficeContentSafetyRisk.Informational);
            cleanupCapability = OfficeContentCleanupCapability.ReportOnly;
            return true;
        }
        XElement? unsupportedAncestor = candidate.ComputedElement.Ancestors().FirstOrDefault(element =>
            element.Name.Namespace != candidate.ComputedElement.Name.Namespace ||
            !IsSupportedSvgTextAncestor(element.Name.LocalName));
        if (unsupportedAncestor != null) {
            bool foreignNamespace = unsupportedAncestor.Name.Namespace != candidate.ComputedElement.Name.Namespace;
            concealment = new SvgContentSafetyConcealment(
                OfficeContentConcealmentKind.NonPrimaryContent,
                "SVG text is nested under a " + (foreignNamespace ? "foreign-namespace " : "unrecognized ") +
                unsupportedAncestor.Name.LocalName +
                " container and is report-only because that ancestor is outside the bounded native paint model.",
                OfficeContentSafetyRisk.Informational);
            cleanupCapability = OfficeContentCleanupCapability.ReportOnly;
            return true;
        }
        if (owner == null && !IsSvgPaintedTextContainer(candidate.ComputedElement)) {
            concealment = new SvgContentSafetyConcealment(
                OfficeContentConcealmentKind.NonPrimaryContent,
                "Text stored directly in a native SVG element outside the supported painted-text structure is machine-readable and report-only.",
                OfficeContentSafetyRisk.Informational);
            cleanupCapability = OfficeContentCleanupCapability.ReportOnly;
            return true;
        }
        if (owner == null) {
            concealment = default;
            cleanupCapability = OfficeContentCleanupCapability.ReportOnly;
            return false;
        }

        string ownerName = owner.Name.LocalName;
        cleanupCapability = ownerName is "script" or "style" or "metadata"
            ? OfficeContentCleanupCapability.ReportOnly
            : OfficeContentCleanupCapability.RemoveText;
        string evidence = ownerName switch {
            "title" => "SVG title text is machine-readable accessibility content but is not painted as ordinary canvas text.",
            "desc" => "SVG description text is machine-readable accessibility content but is not painted as ordinary canvas text.",
            "script" => "SVG script text is machine-readable source content but is not painted as ordinary canvas text.",
            "style" => "SVG stylesheet text is machine-readable source content; removing it could change unrelated rendering and is therefore report-only.",
            _ => "SVG metadata text is machine-readable package content; removing it could invalidate provenance or unrelated metadata and is therefore report-only."
        };
        concealment = new SvgContentSafetyConcealment(
            OfficeContentConcealmentKind.NonPrimaryContent,
            evidence,
            OfficeContentSafetyRisk.Informational);
        return true;
    }

    private static bool IsSvgPaintedTextContainer(XElement element) {
        string name = element.Name.LocalName;
        if (name == "text") return true;
        if (name is not "tspan" and not "textPath" and not "a") return false;
        return element.Ancestors().Any(ancestor => {
            string ancestorName = ancestor.Name.LocalName;
            return ancestorName is "text" or "tspan" or "textPath" or "a";
        });
    }

    private static bool IsSupportedSvgTextAncestor(string name) => name is
        "svg" or "g" or "a" or "switch" or "use" or "defs" or "symbol" or
        "text" or "tspan" or "textPath" or "title" or "desc" or "script" or "style" or "metadata";

    private static ISet<string> CollectSvgReusableTextReferencedIds(XElement root) {
        var ids = new HashSet<string>(StringComparer.Ordinal);
        XNamespace svgNamespace = root.Name.Namespace;
        foreach (XElement use in root.DescendantsAndSelf().Where(element =>
                     IsNativeSvgElement(element, svgNamespace) &&
                     (element.Name.LocalName.Equals("use", StringComparison.Ordinal) ||
                      element.Name.LocalName.Equals("tref", StringComparison.Ordinal)))) {
             foreach (XAttribute href in use.Attributes().Where(attribute =>
                          attribute.Name.LocalName.Equals("href", StringComparison.Ordinal) &&
                         (attribute.Name.NamespaceName.Length == 0 ||
                          attribute.Name.NamespaceName.Equals("http://www.w3.org/1999/xlink", StringComparison.Ordinal)))) {
                string value = TrimSvgCssWhitespace(href.Value);
                if (value.Length > 1 && value[0] == '#' &&
                    value.IndexOfAny(new[] { ' ', '\t', '\r', '\n', '#', '(', ')' }, 1) < 0) {
                    try {
                        string decoded = Uri.UnescapeDataString(value.Substring(1));
                        if (decoded.Length > 0) ids.Add(decoded);
                    } catch (UriFormatException) {
                        ids.Add(value.Substring(1));
                    }
                }
            }
        }
        return ids;
    }

    private static bool HasSvgDynamicRendering(XElement root) {
        XNamespace svgNamespace = root.Name.Namespace;
        return root.DescendantsAndSelf().Any(element => {
            string name = element.Name.LocalName.ToLowerInvariant();
            if (name == "script") return true;
            if (IsNativeSvgElement(element, svgNamespace) &&
                name is "animate" or "animatemotion" or "animatetransform" or "set" or "discard") {
                return true;
            }
            return element.Attributes().Any(attribute => {
                string attributeName = attribute.Name.LocalName;
                if (attribute.Name.NamespaceName.Length == 0 &&
                    attributeName.Length > 2 &&
                    attributeName.StartsWith("on", StringComparison.OrdinalIgnoreCase)) return true;
                return attributeName.Equals("href", StringComparison.OrdinalIgnoreCase) &&
                    IsSvgExecutableUrl(attribute.Value);
            });
        });
    }

    private static bool HasSvgConditionalRendering(XElement root) {
        XNamespace svgNamespace = root.Name.Namespace;
        return root.DescendantsAndSelf().Any(element =>
            IsNativeSvgElement(element, svgNamespace) &&
            element.Attributes().Any(attribute =>
                attribute.Name.NamespaceName.Length == 0 &&
                attribute.Name.LocalName is "systemLanguage" or "requiredFeatures" or "requiredExtensions"));
    }

    private static bool IsSvgExecutableUrl(string value) {
        string normalized = new string(value.Where(character => character is not '\t' and not '\n' and not '\r').ToArray());
        int start = 0;
        while (start < normalized.Length && normalized[start] <= ' ') start++;
        int end = normalized.Length;
        while (end > start && normalized[end - 1] <= ' ') end--;
        return normalized.Substring(start, end - start).StartsWith("javascript:", StringComparison.OrdinalIgnoreCase);
    }

    private static bool TryDescribeSvgContextDependentText(
        XElement element,
        ISet<string> reusableTextIds,
        bool hasDynamicRendering,
        bool hasConditionalRendering,
        out string evidence) {
        XElement logicalOwner = element.AncestorsAndSelf().LastOrDefault(ancestor =>
            ancestor.Name.LocalName.Equals("text", StringComparison.Ordinal)) ?? element;
        if (HasSharedSvgTextPositioning(logicalOwner)) {
            evidence = "SVG text shares character-indexed positioning lists with sibling text nodes and is therefore report-only.";
            return true;
        }
        foreach (XElement current in element.AncestorsAndSelf()) {
            string name = current.Name.LocalName.ToLowerInvariant();
            if (current.Attribute(XNamespace.Xml + "base") != null) {
                evidence = "SVG xml:base changes fragment-reference resolution outside the bounded native projection and is therefore report-only.";
                return true;
            }
            if (current != element) {
                string? opacity = ReadPresentationProperty(current, "opacity")?.Trim();
                if (!string.IsNullOrWhiteSpace(opacity) &&
                    TryUnit(opacity!, out double resolvedOpacity) &&
                    resolvedOpacity < 1D) {
                    evidence = "SVG ancestor opacity is applied after group compositing in browsers and is therefore report-only.";
                    return true;
                }
            }
            if (current.Attribute("textLength") != null || current.Attribute("lengthAdjust") != null) {
                evidence = "SVG text-length adjustment can change browser glyph bounds and is therefore report-only.";
                return true;
            }
            string? textDecoration = ReadPresentationProperty(current, "text-decoration")?.Trim();
            if (!string.IsNullOrWhiteSpace(textDecoration) &&
                !textDecoration!.Equals("none", StringComparison.OrdinalIgnoreCase) &&
                !textDecoration.Equals("initial", StringComparison.OrdinalIgnoreCase) &&
                !textDecoration.Equals("unset", StringComparison.OrdinalIgnoreCase)) {
                evidence = "SVG text decoration paints outside the bounded native glyph projection and is therefore report-only.";
                return true;
            }
            if (TryFindUnmodeledSvgTextGeometryProperty(current, out string unmodeledGeometry)) {
                evidence = "SVG text geometry uses the unmodeled " + unmodeledGeometry +
                    " property and is therefore report-only.";
                return true;
            }
            string? transform = ReadPresentationProperty(current, "transform");
            if (!string.IsNullOrWhiteSpace(transform) &&
                (transform!.Any(character => char.IsWhiteSpace(character) && !IsSvgCssWhitespace(character)) ||
                 !OfficeSvgTransformParser.TryParse(transform, out _))) {
                evidence = "SVG transform syntax is outside the bounded browser grammar and is therefore report-only.";
                return true;
            }
            string? clipPath = ReadPresentationProperty(current, "clip-path")?.Trim();
            if (!string.IsNullOrWhiteSpace(clipPath) &&
                !clipPath!.Equals("none", StringComparison.OrdinalIgnoreCase) &&
                !TryReadBoundedSvgLocalUrlReference(clipPath, out _)) {
                evidence = "SVG clip-path syntax is outside the bounded local URL grammar and is therefore report-only.";
                return true;
            }
            if (!string.IsNullOrWhiteSpace(clipPath) &&
                !clipPath!.Equals("none", StringComparison.OrdinalIgnoreCase) &&
                TryDescribeUnsupportedSvgClipProjection(current, clipPath, out evidence)) {
                return true;
            }
            XAttribute? fontSize = current.Attribute("font-size");
            if (fontSize != null && !IsSvgCssWideKeyword(fontSize.Value) &&
                fontSize.Value.IndexOf("var(", StringComparison.OrdinalIgnoreCase) < 0 &&
                (!TrySvgLength(fontSize.Value, out double resolvedFontSize) || resolvedFontSize <= 0D)) {
                evidence = "SVG text uses font-size syntax outside the bounded native layout subset and is therefore report-only.";
                return true;
            }
            if (HasUnsupportedSvgPresentationPaint(current.Attribute("fill")?.Value) ||
                HasUnsupportedSvgPresentationPaint(current.Attribute("stroke")?.Value)) {
                evidence = "SVG text uses fallback or malformed paint syntax outside the bounded native paint subset and is therefore report-only.";
                return true;
            }
            string? filter = ReadPresentationProperty(current, "filter")?.Trim();
            if (!string.IsNullOrWhiteSpace(filter) &&
                !filter!.Equals("none", StringComparison.OrdinalIgnoreCase)) {
                evidence = "SVG filter output can alter the painted text geometry and is therefore report-only.";
                return true;
            }
            string? mask = ReadPresentationProperty(current, "mask");
            if (!string.IsNullOrWhiteSpace(mask) &&
                !TrimSvgCssWhitespace(mask!).Equals("none", StringComparison.OrdinalIgnoreCase)) {
                evidence = "SVG mask compositing can alter or suppress painted text outside the bounded native projection and is therefore report-only.";
                return true;
            }
            if (name == "svg" && current.Parent != null) {
                string? overflow = ReadPresentationProperty(current, "overflow")?.Trim();
                if (!string.IsNullOrWhiteSpace(overflow) &&
                    !overflow!.Equals("hidden", StringComparison.OrdinalIgnoreCase)) {
                    evidence = "Text inside a nested SVG viewport with non-hidden or inherited overflow depends on browser viewport painting and is therefore report-only.";
                    return true;
                }
            }
            if (name == "switch") {
                evidence = "SVG switch-branch text depends on renderer language and feature context and is therefore report-only.";
                return true;
            }
            if (current.Attributes().Any(attribute =>
                    attribute.Name.NamespaceName.Length == 0 &&
                    (attribute.Name.LocalName.Equals("systemLanguage", StringComparison.Ordinal) ||
                     attribute.Name.LocalName.Equals("requiredFeatures", StringComparison.Ordinal) ||
                     attribute.Name.LocalName.Equals("requiredExtensions", StringComparison.Ordinal)))) {
                evidence = "SVG conditional-processing text depends on renderer language and feature context and is therefore report-only.";
                return true;
            }
            if (name is "defs" or "symbol" or "clippath" or "mask" or "pattern" or "marker" or "filter") {
                evidence = "SVG definition text can affect reusable or composited rendering and is therefore report-only.";
                return true;
            }
            string? id = current.Attribute("id")?.Value;
            if (!string.IsNullOrEmpty(id) && reusableTextIds.Contains(id!)) {
                evidence = "SVG text belongs to a reusable element referenced by use or tref and is therefore report-only.";
                return true;
            }
        }
        if (hasConditionalRendering) {
            evidence = "SVG conditional-processing attributes elsewhere in the document can change browser paint and are therefore report-only for static inspection.";
            return true;
        }
        if (hasDynamicRendering) {
            evidence = "SVG script or animation can change text visibility over time and is therefore report-only for static inspection.";
            return true;
        }
        evidence = string.Empty;
        return false;
    }

    private static bool TryDescribeUnsupportedSvgClipProjection(
        XElement element,
        string clipPath,
        out string evidence) {
        evidence = string.Empty;
        if (!TryReadBoundedSvgLocalUrlReference(clipPath, out string reference)) return false;
        string id;
        try {
            id = Uri.UnescapeDataString(reference.Substring(1));
        } catch (UriFormatException) {
            evidence = "SVG clip-path reference encoding is outside the bounded native clip model and is therefore report-only.";
            return true;
        }
        XElement root = element.AncestorsAndSelf().Last();
        XElement[] definitions = root.DescendantsAndSelf().Where(candidate =>
            candidate.Name.Namespace == root.Name.Namespace &&
            candidate.Name.LocalName.Equals("clipPath", StringComparison.Ordinal) &&
            string.Equals(candidate.Attribute("id")?.Value, id, StringComparison.Ordinal)).Take(2).ToArray();
        if (definitions.Length == 0) return false;
        if (definitions.Length > 1) {
            evidence = "SVG clip-path does not resolve to one native definition and is therefore report-only.";
            return true;
        }
        XElement definition = definitions[0];
        string? nestedClip = ReadPresentationProperty(definition, "clip-path")?.Trim();
        string? units = definition.Attribute("clipPathUnits")?.Value;
        XElement[] geometry = definition.Elements().Where(child =>
            child.Name.Namespace == root.Name.Namespace &&
            child.Name.LocalName is not "title" and not "desc" and not "metadata").Take(2).ToArray();
        string? childClip = geometry.Length == 1 ? ReadPresentationProperty(geometry[0], "clip-path")?.Trim() : null;
        bool unsupported =
            (!string.IsNullOrWhiteSpace(nestedClip) && !nestedClip!.Equals("none", StringComparison.OrdinalIgnoreCase)) ||
            (!string.IsNullOrWhiteSpace(units) && !units!.Equals("userSpaceOnUse", StringComparison.Ordinal)) ||
            geometry.Length > 1 ||
            (geometry.Length == 1 &&
             (geometry[0].Name.LocalName is not "rect" and not "circle" and not "ellipse" and not "polygon" and not "path" ||
              !string.IsNullOrWhiteSpace(childClip) && !childClip!.Equals("none", StringComparison.OrdinalIgnoreCase)));
        if (!unsupported) return false;
        evidence = "SVG clip-path uses compound or otherwise unmodeled clip geometry and is therefore report-only.";
        return true;
    }

    private static bool HasSharedSvgTextPositioning(XElement logicalOwner) {
        if (!logicalOwner.Name.LocalName.Equals("text", StringComparison.Ordinal)) return false;
        return logicalOwner.DescendantsAndSelf().Any(element =>
            new[] { "x", "y", "dx", "dy", "rotate" }.Any(name => {
                string? value = element.Attribute(name)?.Value;
                if (string.IsNullOrWhiteSpace(value)) return false;
                return value!.Split(new[] { ' ', '\t', '\r', '\n', ',' }, StringSplitOptions.RemoveEmptyEntries).Length > 1;
            }));
    }

    private static bool HasUnsupportedSvgPresentationPaint(string? value) {
        if (string.IsNullOrWhiteSpace(value)) return false;
        string normalized = TrimSvgCssWhitespace(value!);
        if (!normalized.StartsWith("url(", StringComparison.OrdinalIgnoreCase)) return false;
        int close = FindSvgCssBlockEnd(normalized, 4, '(', ')');
        bool hasFallback = close >= 0 && close < normalized.Length - 1 &&
            TrimSvgCssWhitespace(normalized.Substring(close + 1)).Length > 0;
        return hasFallback || !TryReadBoundedSvgLocalUrlReference(normalized, out _);
    }
}
