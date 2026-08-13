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
        ref int commandCount) {
        SvgElementReferenceEntryResult result = references.TryEnterDetailed(
            element,
            expectedTargetName,
            out string referenceId,
            out XElement? target);
        if (result is SvgElementReferenceEntryResult.DepthExceeded or SvgElementReferenceEntryResult.Cycle) return false;
        if (result != SvgElementReferenceEntryResult.Entered) {
            string? href = element.Attributes()
                .FirstOrDefault(attribute => attribute.Name.LocalName.Equals("href", StringComparison.OrdinalIgnoreCase))?.Value;
            return string.IsNullOrWhiteSpace(href) || !href!.TrimStart().StartsWith("#", StringComparison.Ordinal);
        }
        try {
            return TryAddRenderedSvgExpansion(
                target!,
                references,
                maximumElements,
                ref elementCount,
                ref commandCount);
        } finally {
            references.Exit(referenceId);
        }
    }
}
