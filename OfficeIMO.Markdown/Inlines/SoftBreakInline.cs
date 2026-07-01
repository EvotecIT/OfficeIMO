namespace OfficeIMO.Markdown;

/// <summary>
/// Soft line break inline. Markdown and HTML render as a newline without forcing a hard break.
/// </summary>
public sealed class SoftBreakInline : MarkdownInline, IRenderableMarkdownInline, IPlainTextMarkdownInline {
    internal string RenderMarkdown() => "\n";
    internal string RenderHtml() => "\n";
    string IRenderableMarkdownInline.RenderMarkdown() => RenderMarkdown();
    string IRenderableMarkdownInline.RenderHtml() => RenderHtml();
    void IPlainTextMarkdownInline.AppendPlainText(System.Text.StringBuilder sb) => sb.Append('\n');
}
