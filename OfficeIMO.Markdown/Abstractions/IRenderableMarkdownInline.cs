namespace OfficeIMO.Markdown;

/// <summary>
/// Rendering contract for inline nodes that can serialize themselves back to Markdown and HTML.
/// Custom inline parser extensions should implement this so the reader, writer, and syntax tree can
/// preserve the node without falling back to opaque object output.
/// </summary>
public interface IRenderableMarkdownInline {
    /// <summary>
    /// Renders the inline node back to Markdown.
    /// </summary>
    string RenderMarkdown();

    /// <summary>
    /// Renders the inline node to HTML.
    /// </summary>
    string RenderHtml();
}
