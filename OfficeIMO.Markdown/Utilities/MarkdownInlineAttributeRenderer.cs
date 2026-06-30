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

        html = RenderNestedStrongEmphasisAttributes(inline, markdownObject.Attributes, html, options);

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

    private static string RenderNestedStrongEmphasisAttributes(
        IMarkdownInline inline,
        MarkdownAttributeSet attributes,
        string html,
        HtmlOptions? options) {
        if (inline is ItalicSequenceInline italic &&
            italic.Inlines.Nodes.Count == 1 &&
            italic.Inlines.Nodes[0] is BoldSequenceInline bold &&
            bold.Attributes.IsEmpty) {
            return InsertAttributesIntoFirstTag(html, "<strong", attributes, options);
        }

        if (inline is BoldSequenceInline strong &&
            strong.Inlines.Nodes.Count == 1 &&
            strong.Inlines.Nodes[0] is ItalicSequenceInline emphasis &&
            emphasis.Attributes.IsEmpty) {
            return InsertAttributesIntoFirstTag(html, "<em", attributes, options);
        }

        return html;
    }

    private static string InsertAttributesIntoFirstTag(string html, string tagPrefix, MarkdownAttributeSet attributes, HtmlOptions? options) {
        var start = html.IndexOf(tagPrefix, StringComparison.OrdinalIgnoreCase);
        if (start < 0) {
            return html;
        }

        var tagEnd = html.IndexOf('>', start);
        if (tagEnd <= start) {
            return html;
        }

        var insertAt = tagEnd;
        if (html[tagEnd - 1] == '/') {
            insertAt = tagEnd - 1;
            if (insertAt > start && html[insertAt - 1] == ' ') {
                insertAt--;
            }
        }

        return html.Substring(0, insertAt) + MarkdownHtmlAttributes.Render(attributes, options) + html.Substring(insertAt);
    }
}
