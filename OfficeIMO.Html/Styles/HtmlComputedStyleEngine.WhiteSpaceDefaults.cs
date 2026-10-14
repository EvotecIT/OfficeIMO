using AngleSharp.Dom;
using OfficeIMO.Html.Css;

namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    /// <summary>Materializes HTML no-break defaults before descendants inherit computed values.</summary>
    private static void ApplyHtmlWhiteSpaceDefaults(
        IElement element,
        bool hasLocalDeclaration,
        IDictionary<string, string> properties,
        ISet<string> inheritedProperties,
        ISet<string> resetProperties,
        ISet<string> originRevertedProperties,
        IDictionary<string, HtmlCssCascadePriority> cascadePriorities,
        IDictionary<string, HtmlCssCascadeTrace>? cascadeTraces) {
        if (element.NamespaceUri != "http://www.w3.org/1999/xhtml") return;
        string? defaultValue = element.LocalName == "nobr" ? "nowrap"
            : element.LocalName == "wbr" && HasHtmlNoBreakAncestor(element) ? "normal"
            : null;
        if (defaultValue == null || hasLocalDeclaration && !originRevertedProperties.Contains("white-space")) return;

        // A local author value, including explicit inheritance or an initial reset,
        // wins. Implicit parent inheritance and author-origin rollback expose the UA default.
        properties["white-space"] = defaultValue;
        inheritedProperties.Remove("white-space");
        resetProperties.Remove("white-space");
        cascadePriorities.Remove("white-space");
        if (cascadeTraces != null && cascadeTraces.TryGetValue("white-space", out HtmlCssCascadeTrace? trace)) {
            cascadeTraces["white-space"] = new HtmlCssCascadeTrace("white-space", defaultValue,
                false, false, trace.Candidates);
        }
    }

    private static bool HasHtmlNoBreakAncestor(IElement element) {
        for (IElement? ancestor = element.ParentElement; ancestor != null; ancestor = ancestor.ParentElement) {
            if (ancestor.NamespaceUri == "http://www.w3.org/1999/xhtml" && ancestor.LocalName == "nobr") return true;
        }
        return false;
    }
}
