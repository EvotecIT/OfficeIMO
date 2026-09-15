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
                string value = href.Value.Trim();
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
        out string evidence) {
        foreach (XElement current in element.AncestorsAndSelf()) {
            string name = current.Name.LocalName.ToLowerInvariant();
            string? transform = ReadPresentationProperty(current, "transform");
            if (!string.IsNullOrWhiteSpace(transform) &&
                transform!.Any(character => char.IsWhiteSpace(character) && !IsSvgCssWhitespace(character))) {
                evidence = "SVG transform syntax contains non-CSS whitespace outside the browser grammar and is therefore report-only.";
                return true;
            }
            XAttribute? fontSize = current.Attribute("font-size");
            if (fontSize != null && !IsSvgCssWideKeyword(fontSize.Value) &&
                fontSize.Value.IndexOf("var(", StringComparison.OrdinalIgnoreCase) < 0 &&
                (!TrySvgLength(fontSize.Value, out double resolvedFontSize) || resolvedFontSize <= 0D)) {
                evidence = "SVG text uses font-size syntax outside the bounded native layout subset and is therefore report-only.";
                return true;
            }
            if (HasSvgFallbackPaint(current.Attribute("fill")?.Value) ||
                HasSvgFallbackPaint(current.Attribute("stroke")?.Value)) {
                evidence = "SVG text uses fallback paint syntax outside the bounded native paint subset and is therefore report-only.";
                return true;
            }
            string? filter = ReadPresentationProperty(current, "filter")?.Trim();
            if (!string.IsNullOrWhiteSpace(filter) &&
                !filter!.Equals("none", StringComparison.OrdinalIgnoreCase)) {
                evidence = "SVG filter output can alter the painted text geometry and is therefore report-only.";
                return true;
            }
            if (name == "svg" && current.Parent != null &&
                string.Equals(ReadPresentationProperty(current, "overflow")?.Trim(), "visible", StringComparison.OrdinalIgnoreCase)) {
                evidence = "Text inside a nested SVG viewport with visible overflow depends on browser viewport painting and is therefore report-only.";
                return true;
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
        if (hasDynamicRendering) {
            evidence = "SVG script or animation can change text visibility over time and is therefore report-only for static inspection.";
            return true;
        }
        evidence = string.Empty;
        return false;
    }

    private static bool HasSvgFallbackPaint(string? value) {
        if (string.IsNullOrWhiteSpace(value)) return false;
        string normalized = value!.Trim();
        if (!normalized.StartsWith("url(", StringComparison.OrdinalIgnoreCase)) return false;
        int close = FindSvgCssBlockEnd(normalized, 4, '(', ')');
        return close >= 0 && close < normalized.Length - 1 &&
            normalized.Substring(close + 1).Trim().Length > 0;
    }
}
