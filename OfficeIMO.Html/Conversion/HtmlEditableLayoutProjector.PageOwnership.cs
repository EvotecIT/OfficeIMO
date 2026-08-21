using AngleSharp.Dom;
using AngleSharp.Html.Dom;

namespace OfficeIMO.Html;

public static partial class HtmlEditableLayoutProjector {
    private static IReadOnlyList<double> CreateForcedPageBreakBoundaries(
        IReadOnlyList<IElement> elements,
        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles) {
        var boundaries = new List<double>();
        for (int index = 0; index < elements.Count; index++) {
            if (!styles.TryGetValue(elements[index], out HtmlComputedStyle? style)) continue;
            if (IsForcedPageBreakValue(style.GetValue("break-before"))
                || IsForcedPageBreakValue(style.GetValue("page-break-before"))) {
                boundaries.Add(index - 0.25D);
            }
            if (IsForcedPageBreakValue(style.GetValue("break-after"))
                || IsForcedPageBreakValue(style.GetValue("page-break-after"))) {
                boundaries.Add(index + 0.25D);
            }
        }
        return boundaries.AsReadOnly();
    }

    private static bool TryGetForcedPageBreakOwnershipDetail(
        int elementOrder,
        IReadOnlyList<double> forcedPageBreakBoundaries,
        bool preserveRegionsBeforeForcedPageBreaks,
        bool preserveRegionsAfterForcedPageBreaks,
        out string detail) {
        if (preserveRegionsBeforeForcedPageBreaks
            && forcedPageBreakBoundaries.Any(boundary => boundary > elementOrder)) {
            detail = "forcedPageBreakAfter=true";
            return true;
        }
        if (preserveRegionsAfterForcedPageBreaks
            && forcedPageBreakBoundaries.Any(boundary => boundary < elementOrder)) {
            detail = "forcedPageBreakBefore=true";
            return true;
        }
        detail = string.Empty;
        return false;
    }

    private static bool IsForcedPageBreakValue(string? value) {
        string normalized = (value ?? string.Empty).Trim().ToLowerInvariant();
        return normalized == "page" || normalized == "always"
            || normalized == "left" || normalized == "right"
            || normalized == "recto" || normalized == "verso";
    }

    private static bool ContainsUnrenderedSourceImage(
        HtmlRenderLayoutRegion region,
        IElement regionElement,
        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles) {
        var renderedKeys = new HashSet<string>(
            EnumerateImages(region.Visuals, includeBackgroundImages: false)
                .Select(item => item.Image.Source)
                .Where(source => !string.IsNullOrWhiteSpace(source))
                .Cast<string>(),
            StringComparer.Ordinal);
        foreach (IHtmlImageElement image in regionElement.QuerySelectorAll("img").OfType<IHtmlImageElement>()) {
            if (!IsProjectionImageVisible(image, regionElement, styles)) continue;
            string? imageKey = GetImageSourceKey(image);
            if (!string.IsNullOrWhiteSpace(imageKey)
                && !renderedKeys.Contains(DescribeImageSource(imageKey))) {
                return true;
            }
        }
        return false;
    }
}