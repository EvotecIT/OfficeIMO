namespace OfficeIMO.MarkdownRenderer;

/// <summary>
/// Pre-processes Markdown text before it is parsed by <see cref="MarkdownRenderer.RenderBodyHtml"/>.
/// </summary>
/// <param name="markdown">Input markdown text.</param>
/// <param name="options">Active renderer options.</param>
/// <returns>Transformed markdown text.</returns>
public delegate string MarkdownTextPreProcessor(string markdown, MarkdownRendererOptions options);
