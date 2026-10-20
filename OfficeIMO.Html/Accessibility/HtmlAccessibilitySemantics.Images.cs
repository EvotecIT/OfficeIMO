using AngleSharp.Dom;

namespace OfficeIMO.Html;

public static partial class HtmlAccessibilitySemantics {
    /// <summary>Determines whether image paint is explicitly excluded from accessible content.</summary>
    internal static bool IsDecorativeImage(IElement element) {
        // An image input is an action and still requires an accessible name.
        if (element.TagName.Equals("INPUT", StringComparison.OrdinalIgnoreCase)) return false;
        return IsAriaHidden(element)
            || ContainsToken(element.GetAttribute("role"), "none")
            || ContainsToken(element.GetAttribute("role"), "presentation")
            || element.TagName.Equals("IMG", StringComparison.OrdinalIgnoreCase)
                && element.HasAttribute("alt") && string.IsNullOrWhiteSpace(element.GetAttribute("alt"))
                && string.IsNullOrWhiteSpace(GetImageAccessibleName(element));
    }
}
