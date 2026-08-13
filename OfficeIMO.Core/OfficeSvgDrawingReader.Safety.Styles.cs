using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    private static string? ReadRasterPresentationProperty(XElement element, string propertyName) {
        string? value = null;
        foreach (XAttribute attribute in element.Attributes()) {
            if (attribute.IsNamespaceDeclaration || attribute.Name.Namespace == XNamespace.Xml) continue;
            if (attribute.Name.LocalName.Equals(propertyName, StringComparison.Ordinal)) value = attribute.Value;
        }

        string? style = ReadRasterInlineStyleAttribute(element);
        if (string.IsNullOrWhiteSpace(style)) return value;
        int selectedPriority = -1;
        foreach (string declaration in SplitRasterStyleDeclarations(StripCssComments(style!))) {
            int colon = declaration.IndexOf(':');
            if (colon <= 0 || !declaration.Substring(0, colon).Trim().Equals(propertyName, StringComparison.Ordinal)) continue;
            string candidate = NormalizeInlineStyleValue(declaration.Substring(colon + 1), out int priority);
            if (priority < selectedPriority) continue;
            value = candidate;
            selectedPriority = priority;
        }
        return value;
    }

    private static IEnumerable<string> SplitRasterStyleDeclarations(string style) {
        int start = 0;
        int depth = 0;
        char quote = '\0';
        for (int index = 0; index < style.Length; index++) {
            char current = style[index];
            if (quote != '\0') {
                if (current == '\\') {
                    index++;
                } else if (current == quote) {
                    quote = '\0';
                }
                continue;
            }
            if (current is '\'' or '"') {
                quote = current;
            } else if (current == '(') {
                depth++;
            } else if (current == ')' && depth > 0) {
                depth--;
            } else if (current == ';' && depth == 0) {
                string declaration = style.Substring(start, index - start).Trim();
                if (declaration.Length > 0) yield return declaration;
                start = index + 1;
            }
        }
        string tail = style.Substring(start).Trim();
        if (tail.Length > 0) yield return tail;
    }

    private static bool TryAddRenderedSvgDefinitionExpansion(
        XElement target,
        SvgElementReferenceRegistry references,
        int maximumElements,
        ref int elementCount,
        ref int commandCount) {
        ResolveRenderedSvgAncestorPaint(
            target,
            out string? fill,
            out string? stroke,
            out string? markerStart,
            out string? markerMid,
            out string? markerEnd);
        return TryAddRenderedSvgExpansion(
            target,
            references,
            maximumElements,
            ref elementCount,
            ref commandCount,
            fill,
            stroke,
            markerStart,
            markerMid,
            markerEnd);
    }

    private static void ResolveRenderedSvgAncestorPaint(
        XElement target,
        out string? fill,
        out string? stroke,
        out string? markerStart,
        out string? markerMid,
        out string? markerEnd) {
        fill = stroke = markerStart = markerMid = markerEnd = null;
        foreach (XElement ancestor in target.Ancestors().Reverse()) {
            fill = ResolveInheritedSvgPaint(ancestor, "fill", fill);
            stroke = ResolveInheritedSvgPaint(ancestor, "stroke", stroke);
            string? marker = ResolveInheritedSvgPaint(ancestor, "marker", inherited: null);
            markerStart = ResolveInheritedSvgPaint(ancestor, "marker-start", marker ?? markerStart);
            markerMid = ResolveInheritedSvgPaint(ancestor, "marker-mid", marker ?? markerMid);
            markerEnd = ResolveInheritedSvgPaint(ancestor, "marker-end", marker ?? markerEnd);
        }
    }
}
