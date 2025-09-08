namespace OfficeIMO.Markdown;

/// <summary>
/// Represents a top-level Markdown block that can render to Markdown and HTML.
/// </summary>
public interface IMarkdownBlock {
    /// <summary>Renders the block as Markdown text.</summary>
    string RenderMarkdown();
    /// <summary>Renders a simple HTML representation of the block.</summary>
    string RenderHtml();
}
