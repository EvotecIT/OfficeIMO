using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using System.Threading;

namespace OfficeIMO.Html;

/// <summary>Inspects source image names using the operation's resolved CSS visibility.</summary>
internal static class HtmlAccessibilitySourceInspection {
    internal static int CountImagesMissingNames(IHtmlDocument document,
        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles, CancellationToken cancellationToken = default) {
        int count = 0;
        foreach (IElement element in document.QuerySelectorAll("img,input[type=image],svg")) {
            cancellationToken.ThrowIfCancellationRequested();
            if (IsExcluded(element, styles)) continue;
            if (string.IsNullOrWhiteSpace(HtmlAccessibilitySemantics.GetImageAccessibleName(element))) count++;
        }
        return count;
    }

    private static bool IsExcluded(IElement element, IReadOnlyDictionary<IElement, HtmlComputedStyle> styles) {
        bool isImageInput = element.TagName.Equals("INPUT", StringComparison.OrdinalIgnoreCase);
        if (HtmlAccessibilitySemantics.IsDecorativeImage(element)) return true;
        if (styles.TryGetValue(element, out HtmlComputedStyle? elementStyle)) {
            string visibility = elementStyle.GetValue("visibility").Trim();
            if (visibility.Equals("hidden", StringComparison.OrdinalIgnoreCase) ||
                visibility.Equals("collapse", StringComparison.OrdinalIgnoreCase)) return true;
        }
        for (IElement? current = element; current != null; current = current.ParentElement) {
            if (current.HasAttribute("hidden") || !isImageInput && HtmlAccessibilitySemantics.IsAriaHidden(current)) return true;
            if (styles.TryGetValue(current, out HtmlComputedStyle? style) &&
                style.GetValue("display").Trim().Equals("none", StringComparison.OrdinalIgnoreCase)) return true;
        }
        return false;
    }
}
