using AngleSharp.Dom;
using OfficeIMO.Html.Css;

namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    /// <summary>Applies the align attribute hint and centers HTML header cells only when their parent retains the initial text alignment.</summary>
    private static void ApplyHtmlTableAlignmentDefault(
        IElement element,
        bool hasLocalDeclaration,
        IReadOnlyDictionary<string, string>? parentProperties,
        IDictionary<string, string> properties,
        ISet<string> inheritedProperties,
        ISet<string> resetProperties,
        ISet<string> originRevertedProperties,
        IDictionary<string, HtmlCssCascadePriority> cascadePriorities,
        IDictionary<string, HtmlCssCascadeTrace>? cascadeTraces) {
        if (element.NamespaceUri != "http://www.w3.org/1999/xhtml") return;
        // The align attribute is a presentational hint below author CSS and above the th default.
        // HTML maps it to text-align for table parts and for the legacy block elements.
        if (element.LocalName is "td" or "th" or "tr" or "thead" or "tbody" or "tfoot"
                or "p" or "div" or "h1" or "h2" or "h3" or "h4" or "h5" or "h6"
            && ResolveTableAlignAttribute(element) is string hint) {
            ApplyHtmlUserAgentDefault("text-align", hint, hasLocalDeclaration, properties,
                inheritedProperties, resetProperties, originRevertedProperties, cascadePriorities, cascadeTraces);
            return;
        }
        if (element.LocalName != "th") return;
        // HTML's conditional UA rule leaves authored parent alignment intact.
        // An omitted or reset value represents CSS's logical initial value, start.
        if (parentProperties != null && parentProperties.TryGetValue("text-align", out string? alignment)
            && !string.IsNullOrWhiteSpace(alignment) && !string.Equals(alignment, "start", StringComparison.OrdinalIgnoreCase)) return;
        ApplyHtmlUserAgentDefault("text-align", "center", hasLocalDeclaration, properties,
            inheritedProperties, resetProperties, originRevertedProperties, cascadePriorities, cascadeTraces);
    }

    private static string? ResolveTableAlignAttribute(IElement element) {
        string? value = element.GetAttribute("align")?.Trim().ToLowerInvariant();
        return value switch {
            "left" or "right" or "center" or "justify" => value,
            "middle" => "center",
            _ => null
        };
    }
}
