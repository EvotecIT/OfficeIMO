using AngleSharp.Dom;

namespace OfficeIMO.Html;

/// <summary>Reads stylesheet text without copying it on every pass.</summary>
internal static class HtmlStyleElementText {
    // AngleSharp concatenates descendant text into a new string for every TextContent read. A style element
    // normally holds one text node, whose data is exactly that text, so large embedded stylesheets
    // (for example @font-face data URIs) are not copied once per stylesheet pass.
    internal static string Read(INode element) =>
        element.FirstChild is IText text && text.NextSibling == null ? text.Data : element.TextContent;
}
