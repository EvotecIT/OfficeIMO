using AngleSharp.Dom;
using OfficeIMO.Html.Css;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    /// <summary>Preserves initial font values when a consumer projects the computed property map.</summary>
    private static void MaterializeInitialFontValues(
        IDictionary<string, string> properties,
        ISet<string> resetProperties,
        ISet<string> originRevertedProperties,
        IDictionary<string, HtmlCssCascadeTrace>? cascadeTraces) {
        foreach (string property in resetProperties) {
            if (property is not ("font-weight" or "font-style") || originRevertedProperties.Contains(property)) continue;
            properties[property] = "normal";
            if (cascadeTraces != null && cascadeTraces.TryGetValue(property, out HtmlCssCascadeTrace? trace)) {
                cascadeTraces[property] = new HtmlCssCascadeTrace(property, "normal",
                    false, true, trace.Candidates);
            }
        }
    }

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
        if (defaultValue == null) return;
        ApplyHtmlUserAgentDefault("white-space", defaultValue, hasLocalDeclaration, properties,
            inheritedProperties, resetProperties, originRevertedProperties, cascadePriorities, cascadeTraces);
    }

    /// <summary>Materializes semantic font defaults before descendants resolve inheritance and relative weights.</summary>
    private static void ApplyHtmlFontDefaults(
        IElement element,
        IDictionary<string, CascadedProperty> declarations,
        IReadOnlyDictionary<string, string>? parentProperties,
        IDictionary<string, string> properties,
        ISet<string> inheritedProperties,
        ISet<string> resetProperties,
        ISet<string> originRevertedProperties,
        IDictionary<string, HtmlCssCascadePriority> cascadePriorities,
        IDictionary<string, HtmlCssCascadeTrace>? cascadeTraces) {
        if (element.NamespaceUri != "http://www.w3.org/1999/xhtml") return;
        string tag = element.LocalName;
        if (tag is "i" or "em" or "cite" or "dfn" or "var" or "address") {
            ApplyHtmlUserAgentDefault("font-style", "italic", declarations.ContainsKey("font-style"), properties,
                inheritedProperties, resetProperties, originRevertedProperties, cascadePriorities, cascadeTraces);
        }
        string? weight = tag is "b" or "strong" ? "bolder"
            : tag is "h1" or "h2" or "h3" or "h4" or "h5" or "h6" or "th" ? "bold"
            : null;
        if (weight == null) return;
        int inheritedWeight = parentProperties != null && parentProperties.TryGetValue("font-weight", out string? parentWeight)
            && OfficeFontFaceCssParser.TryWeight(parentWeight, 400, out int parsedWeight) ? parsedWeight : 400;
        OfficeFontFaceCssParser.TryWeight(weight, inheritedWeight, out int defaultWeight);
        ApplyHtmlUserAgentDefault("font-weight", defaultWeight.ToString(System.Globalization.CultureInfo.InvariantCulture),
            declarations.ContainsKey("font-weight"), properties,
            inheritedProperties, resetProperties, originRevertedProperties, cascadePriorities, cascadeTraces);
    }

    /// <summary>Exposes an HTML default only when no local author value survives the cascade.</summary>
    private static void ApplyHtmlUserAgentDefault(
        string property,
        string defaultValue,
        bool hasLocalDeclaration,
        IDictionary<string, string> properties,
        ISet<string> inheritedProperties,
        ISet<string> resetProperties,
        ISet<string> originRevertedProperties,
        IDictionary<string, HtmlCssCascadePriority> cascadePriorities,
        IDictionary<string, HtmlCssCascadeTrace>? cascadeTraces) {
        if (hasLocalDeclaration && !originRevertedProperties.Contains(property)) return;

        // A local author value, including explicit inheritance or an initial reset,
        // wins. Implicit parent inheritance and author-origin rollback expose the UA default.
        properties[property] = defaultValue;
        inheritedProperties.Remove(property);
        resetProperties.Remove(property);
        cascadePriorities.Remove(property);
        if (cascadeTraces != null && cascadeTraces.TryGetValue(property, out HtmlCssCascadeTrace? trace)) {
            cascadeTraces[property] = new HtmlCssCascadeTrace(property, defaultValue,
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
