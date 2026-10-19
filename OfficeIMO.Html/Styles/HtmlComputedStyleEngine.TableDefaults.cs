using AngleSharp.Dom;
using OfficeIMO.Html.Css;

namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    /// <summary>Centers HTML header cells only when their parent retains the initial text alignment.</summary>
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
        if (element.NamespaceUri != "http://www.w3.org/1999/xhtml" || element.LocalName != "th") return;
        // HTML's conditional UA rule leaves authored parent alignment intact.
        // An omitted or reset value represents CSS's logical initial value, start.
        if (parentProperties != null && parentProperties.TryGetValue("text-align", out string? alignment)
            && !string.IsNullOrWhiteSpace(alignment) && !string.Equals(alignment, "start", StringComparison.OrdinalIgnoreCase)) return;
        ApplyHtmlUserAgentDefault("text-align", "center", hasLocalDeclaration, properties,
            inheritedProperties, resetProperties, originRevertedProperties, cascadePriorities, cascadeTraces);
    }
}
