namespace OfficeIMO.Markdown;

/// <summary>
/// Hard line break inline. Markdown renders as two spaces + newline; HTML as <br/>.
/// </summary>
public sealed class HardBreakInline : IMarkdownInline {
    internal string RenderMarkdown() => "  \n";
    internal string RenderHtml() => "<br/>";
}
