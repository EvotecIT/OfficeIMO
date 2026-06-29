namespace OfficeIMO.Markdown;

internal static class MarkdownInlineAttributeRenderer {
    internal static string RenderMarkdown(IMarkdownInline inline, string markdown) {
        if (inline is not MarkdownObject markdownObject || markdownObject.Attributes.IsEmpty) {
            return markdown;
        }

        return markdown + MarkdownAttributeBlockRenderer.RenderInlineTrailing(markdownObject.Attributes);
    }

    internal static string RenderHtml(IMarkdownInline inline, string html, HtmlOptions? options) {
        if (inline is not MarkdownObject markdownObject || markdownObject.Attributes.IsEmpty || string.IsNullOrEmpty(html)) {
            return html;
        }

        if (html[0] != '<' || (html.Length > 1 && (html[1] == '/' || html[1] == '!' || html[1] == '?'))) {
            return html;
        }

        int tagEnd = html.IndexOf('>');
        if (tagEnd <= 0) {
            return html;
        }

        int insertAt = tagEnd;
        if (html[tagEnd - 1] == '/') {
            insertAt = tagEnd - 1;
            if (insertAt > 0 && html[insertAt - 1] == ' ') {
                insertAt--;
            }
        }

        var attributes = MarkdownHtmlAttributes.Render(markdownObject.Attributes, options);
        return html.Substring(0, insertAt) + attributes + html.Substring(insertAt);
    }
}
