namespace OfficeIMO.Html;

/// <summary>
/// Shared HTML element facts used by OfficeIMO converters that walk AngleSharp DOM nodes.
/// </summary>
public static class HtmlDomElementFacts {
    /// <summary>
    /// Returns whether the specified HTML local name is a void element that cannot have an end tag.
    /// </summary>
    /// <param name="localName">HTML element local name.</param>
    /// <returns><c>true</c> when the element is an HTML void element; otherwise <c>false</c>.</returns>
    public static bool IsVoidElement(string? localName) {
        switch ((localName ?? string.Empty).Trim().ToLowerInvariant()) {
            case "area":
            case "base":
            case "br":
            case "col":
            case "embed":
            case "hr":
            case "img":
            case "input":
            case "link":
            case "meta":
            case "param":
            case "source":
            case "track":
            case "wbr":
                return true;
            default:
                return false;
        }
    }
}
