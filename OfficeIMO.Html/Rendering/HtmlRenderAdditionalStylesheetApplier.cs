using AngleSharp.Html.Dom;

namespace OfficeIMO.Html;

/// <summary>Appends caller-provided author stylesheets after all document content in the cloned render DOM.</summary>
internal static class HtmlRenderAdditionalStylesheetApplier {
    internal static void Apply(IHtmlDocument document, IReadOnlyList<string> stylesheets) {
        if (stylesheets.Count == 0) return;
        IHtmlElement? html = document.DocumentElement as IHtmlElement;
        if (html == null) return;

        foreach (string css in stylesheets) {
            if (string.IsNullOrWhiteSpace(css)) continue;
            IHtmlStyleElement? style = document.CreateElement("style") as IHtmlStyleElement;
            if (style == null) continue;
            style.SetAttribute("data-officeimo-render-stylesheet", "caller");
            style.TextContent = css;
            // CSS consumers enumerate stylesheet nodes in document order. Keeping these
            // nodes after the body makes caller CSS authoritative even when source HTML
            // contains a non-conforming body-positioned style or stylesheet link.
            html.AppendChild(style);
        }
    }
}
