namespace OfficeIMO.Html;

/// <summary>Document-local HTML/SVG ID relationships used during publication projection and validation.</summary>
internal static class HtmlIdentifierReferences {
    private static readonly HashSet<string> AriaReferences = new HashSet<string>(new[] {
        "aria-activedescendant", "aria-controls", "aria-describedby", "aria-details", "aria-errormessage",
        "aria-flowto", "aria-labelledby", "aria-owns"
    }, StringComparer.Ordinal);

    internal static bool IsReference(string namespaceUri, string element, string attribute) {
        bool html = namespaceUri == "http://www.w3.org/1999/xhtml";
        if (!html && namespaceUri != "http://www.w3.org/2000/svg") return false;
        return AriaReferences.Contains(attribute) || html && (attribute == "itemref" ||
            attribute == "headers" && (element == "td" || element == "th") ||
            attribute == "for" && (element == "label" || element == "output") ||
            attribute == "list" && element == "input" ||
            attribute == "form" && new[] { "button", "fieldset", "input", "object", "output", "select", "textarea" }.Contains(element, StringComparer.Ordinal));
    }
}
