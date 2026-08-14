using System;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    private static bool TryAddRenderedSvgElementReference(
        XElement element,
        string expectedTargetName,
        SvgElementReferenceRegistry references,
        int maximumElements,
        ref int elementCount,
        ref int commandCount,
        OfficeTransform transform,
        double viewX,
        double viewY,
        SvgRasterWorkBudget rasterWork) {
        SvgElementReferenceEntryResult result = references.TryEnterDetailed(
            element,
            expectedTargetName,
            out string referenceId,
            out XElement? target);
        if (result is SvgElementReferenceEntryResult.DepthExceeded or SvgElementReferenceEntryResult.Cycle) return false;
        if (result != SvgElementReferenceEntryResult.Entered) return !HasLocalSvgElementReference(element);
        try {
            return TryResolveRenderedSvgAncestorStrokeStyle(target!, out SvgRasterStrokeStyle strokeStyle)
                && TryResolveRenderedSvgAncestorTextStyle(target!, out SvgRasterTextStyle textStyle)
                && TryAddRenderedSvgExpansion(
                target!,
                references,
                maximumElements,
                ref elementCount,
                ref commandCount,
                transform,
                viewX,
                viewY,
                rasterWork,
                inheritedStrokeStyle: strokeStyle,
                inheritedTextStyle: textStyle);
        } finally {
            references.Exit(referenceId);
        }
    }

    private static bool HasLocalSvgElementReference(XElement element) {
        XAttribute[] hrefAttributes = element.Attributes()
            .Where(attribute => attribute.Name.LocalName.Equals("href", StringComparison.OrdinalIgnoreCase))
            .Take(2)
            .ToArray();
        if (hrefAttributes.Length > 1) return true;
        string? href = hrefAttributes.FirstOrDefault()?.Value;
        return !string.IsNullOrWhiteSpace(href) && href!.TrimStart().StartsWith("#", StringComparison.Ordinal);
    }
}
