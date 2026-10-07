using AngleSharp.Dom;

namespace OfficeIMO.Html;

/// <summary>Shared HTML display defaults used after the author cascade.</summary>
internal static class HtmlElementDisplay {
    internal static string GetDefaultValue(IElement element) {
        if (element.HasAttribute("hidden")) return "none";
        string tag = element.TagName.ToLowerInvariant();
        if (tag == "math" && string.Equals(element.GetAttribute("display"), "block", StringComparison.OrdinalIgnoreCase)) return "block";
        if (tag == "dialog" && !element.HasAttribute("open")) return "none";
        if (tag == "li") return "list-item";
        if (tag == "table") return "table";
        if (tag == "caption") return "table-caption";
        if (tag == "colgroup") return "table-column-group";
        if (tag == "col") return "table-column";
        if (tag == "thead") return "table-header-group";
        if (tag == "tbody") return "table-row-group";
        if (tag == "tfoot") return "table-footer-group";
        if (tag == "tr") return "table-row";
        if (tag == "td" || tag == "th") return "table-cell";
        return IsDefaultBlockTag(tag) ? "block" : "inline";
    }

    internal static bool IsDefaultBlockTag(string tagName) {
        string tag = tagName.ToLowerInvariant();
        return tag == "html" || tag == "body" || tag == "address" || tag == "article" || tag == "aside" || tag == "blockquote"
            || tag == "details" || tag == "dialog" || tag == "div" || tag == "dl" || tag == "dt" || tag == "dd" || tag == "fieldset"
            || tag == "figcaption" || tag == "figure" || tag == "footer" || tag == "form" || tag == "h1" || tag == "h2" || tag == "h3"
            || tag == "h4" || tag == "h5" || tag == "h6" || tag == "header" || tag == "hr" || tag == "li" || tag == "main"
            || tag == "marquee" || tag == "nav" || tag == "ol" || tag == "p" || tag == "pre" || tag == "section" || tag == "summary" || tag == "table"
            || tag == "ul";
    }
}
