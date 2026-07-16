using AngleSharp.Dom;

namespace OfficeIMO.Html;

/// <summary>
/// Resolves the small, deterministic subset of HTML and ARIA accessibility semantics
/// used by OfficeIMO document conversion pipelines.
/// </summary>
public static class HtmlAccessibilitySemantics {
    private static readonly char[] TokenSeparators = { ' ', '\t', '\r', '\n', '\f' };

    /// <summary>Returns whether an element declares the requested ARIA role.</summary>
    public static bool HasRole(IElement element, string role) =>
        element != null && ContainsToken(element.GetAttribute("role"), role);

    /// <summary>Returns whether an element declares the requested EPUB structural semantic.</summary>
    public static bool HasEpubType(IElement element, string semanticType) {
        if (element == null || string.IsNullOrWhiteSpace(semanticType)) return false;

        string? value = element.GetAttribute("epub:type");
        if (string.IsNullOrWhiteSpace(value)) {
            foreach (IAttr attribute in element.Attributes) {
                if (attribute.Name.Equals("epub:type", StringComparison.OrdinalIgnoreCase)) {
                    value = attribute.Value;
                    break;
                }
            }
        }
        return ContainsToken(value, semanticType);
    }

    /// <summary>
    /// Resolves a heading level from a native heading element or an ARIA heading role.
    /// ARIA levels outside the Markdown heading range are clamped to levels 1 through 6.
    /// </summary>
    public static bool TryGetHeadingLevel(IElement element, out int level) {
        level = 0;
        if (element == null) return false;

        string tagName = element.TagName;
        if (tagName.Length == 2
            && (tagName[0] == 'h' || tagName[0] == 'H')
            && tagName[1] >= '1'
            && tagName[1] <= '6') {
            level = tagName[1] - '0';
            return true;
        }

        if (!HasRole(element, "heading")) return false;
        if (!int.TryParse(
                element.GetAttribute("aria-level"),
                System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture,
                out level)) {
            level = 2;
        }
        if (level < 1) level = 1;
        if (level > 6) level = 6;
        return true;
    }

    /// <summary>
    /// Resolves an accessible name using ARIA labelling, host-language alternatives,
    /// optional element text, and title fallback. This does not mutate the DOM.
    /// </summary>
    /// <param name="element">Element to name.</param>
    /// <param name="includeTextFallback">Whether normalized descendant text may supply the name.</param>
    public static string GetAccessibleName(IElement element, bool includeTextFallback = false) =>
        GetAccessibleName(element, includeTextFallback, treatAsImage: false);

    /// <summary>
    /// Resolves an accessible image name, including image <c>alt</c> semantics for custom
    /// elements that a converter explicitly aliases to an image.
    /// </summary>
    public static string GetImageAccessibleName(IElement element) =>
        GetAccessibleName(element, includeTextFallback: false, treatAsImage: true);

    private static string GetAccessibleName(IElement element, bool includeTextFallback, bool treatAsImage) {
        if (element == null) return string.Empty;

        string labelledBy = ResolveLabelledBy(element);
        if (labelledBy.Length > 0) return labelledBy;

        string ariaLabel = NormalizeText(element.GetAttribute("aria-label"));
        if (ariaLabel.Length > 0) return ariaLabel;

        string tagName = element.TagName;
        if ((treatAsImage
             || tagName.Equals("IMG", StringComparison.OrdinalIgnoreCase)
             || tagName.Equals("AREA", StringComparison.OrdinalIgnoreCase))
            && element.HasAttribute("alt")) {
            return NormalizeText(element.GetAttribute("alt"));
        }
        if (tagName.Equals("INPUT", StringComparison.OrdinalIgnoreCase)
            && string.Equals(element.GetAttribute("type"), "image", StringComparison.OrdinalIgnoreCase)
            && element.HasAttribute("alt")) {
            return NormalizeText(element.GetAttribute("alt"));
        }
        if (tagName.Equals("SVG", StringComparison.OrdinalIgnoreCase)) {
            IElement? titleElement = element.Children.FirstOrDefault(static child =>
                child.TagName.Equals("TITLE", StringComparison.OrdinalIgnoreCase));
            string svgTitle = NormalizeText(titleElement?.TextContent);
            if (svgTitle.Length > 0) return svgTitle;
        }

        if (includeTextFallback) {
            string text = NormalizeText(element.TextContent);
            if (text.Length > 0) return text;
        }

        return NormalizeText(element.GetAttribute("title"));
    }

    /// <summary>Returns whether an element is explicitly hidden from the accessibility tree.</summary>
    public static bool IsAriaHidden(IElement element) =>
        element != null
        && string.Equals(element.GetAttribute("aria-hidden")?.Trim(), "true", StringComparison.OrdinalIgnoreCase);

    internal static bool ContainsToken(string? value, string token) {
        if (string.IsNullOrWhiteSpace(value) || string.IsNullOrWhiteSpace(token)) return false;
        foreach (string candidate in value!.Split(TokenSeparators, StringSplitOptions.RemoveEmptyEntries)) {
            if (candidate.Equals(token, StringComparison.OrdinalIgnoreCase)) return true;
        }
        return false;
    }

    private static string ResolveLabelledBy(IElement element) {
        string? value = element.GetAttribute("aria-labelledby");
        if (string.IsNullOrWhiteSpace(value) || element.Owner == null) return string.Empty;

        var labels = new List<string>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (string id in value!.Split(TokenSeparators, StringSplitOptions.RemoveEmptyEntries)) {
            if (!seen.Add(id)) continue;
            IElement? label = element.Owner.GetElementById(id);
            if (label == null || ReferenceEquals(label, element)) continue;
            string text = NormalizeText(label.TextContent);
            if (text.Length > 0) labels.Add(text);
        }
        return string.Join(" ", labels);
    }

    private static string NormalizeText(string? value) {
        if (string.IsNullOrWhiteSpace(value)) return string.Empty;
        return string.Join(" ", value!.Split(TokenSeparators, StringSplitOptions.RemoveEmptyEntries));
    }
}
