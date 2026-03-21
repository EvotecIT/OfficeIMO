namespace OfficeIMO.Markdown;

/// <summary>
/// Allows a block to render HTML with access to the surrounding body render context.
/// Custom blocks can implement this when their HTML output depends on <see cref="HtmlOptions"/>
/// or the surrounding block list rather than only the block's own local state.
/// </summary>
public interface IContextualHtmlMarkdownBlock {
    /// <summary>
    /// Renders the block to HTML using the supplied body render context.
    /// </summary>
    /// <param name="context">Current body render context.</param>
    /// <returns>Rendered HTML for the block.</returns>
    string RenderHtml(MarkdownBodyRenderContext context);
}
