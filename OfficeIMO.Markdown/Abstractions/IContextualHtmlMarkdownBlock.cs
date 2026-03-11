namespace OfficeIMO.Markdown;

internal interface IContextualHtmlMarkdownBlock {
    string RenderHtml(MarkdownBodyRenderContext context);
}
