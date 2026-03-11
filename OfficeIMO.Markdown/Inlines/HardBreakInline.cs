namespace OfficeIMO.Markdown;

/// <summary>
/// Hard line break inline. Markdown renders as two spaces + newline; HTML as <br/>.
/// </summary>
public sealed class HardBreakInline : IMarkdownInline, IRenderableMarkdownInline {
    internal string RenderMarkdown() => "  \n";
    internal string RenderHtml() => "<br/>";
    string IRenderableMarkdownInline.RenderMarkdown() => RenderMarkdown();
    string IRenderableMarkdownInline.RenderHtml() => RenderHtml();
}
