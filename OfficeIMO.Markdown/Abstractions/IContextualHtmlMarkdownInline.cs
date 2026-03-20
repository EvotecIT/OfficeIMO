namespace OfficeIMO.Markdown;

/// <summary>
/// Allows an inline node to render HTML with access to the active <see cref="HtmlOptions"/>.
/// Implement this on custom inline nodes when HTML output depends on rendering policies such as
/// title, URL/image policy, theme, or other option-driven behavior rather than only local state.
/// </summary>
public interface IContextualHtmlMarkdownInline {
    /// <summary>
    /// Renders the inline node to HTML using the supplied rendering options.
    /// </summary>
    /// <param name="options">Active HTML rendering options.</param>
    /// <returns>Rendered HTML for the inline.</returns>
    string RenderHtml(HtmlOptions options);
}
