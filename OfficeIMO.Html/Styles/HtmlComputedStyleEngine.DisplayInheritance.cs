using AngleSharp.Dom;

namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    private static IReadOnlyDictionary<string, string>? GetParentPropertiesWithEffectiveDisplay(
        HtmlComputedStyle? parent,
        IElement? parentElement) {
        if (parent == null) return null;
        if (parentElement == null || parent.Properties.ContainsKey("display")) return parent.Properties;

        // The bounded cascade does not materialize user-agent display defaults. CSS
        // display: inherit still needs the parent's computed display, including that default.
        var properties = new Dictionary<string, string>(HtmlCssPropertyNameComparer.Instance);
        foreach (KeyValuePair<string, string> property in parent.Properties) {
            properties.Add(property.Key, property.Value);
        }
        properties["display"] = HtmlRenderStyleResolver.ResolveDisplay(parentElement, string.Empty);
        return properties;
    }
}
