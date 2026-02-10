namespace OfficeIMO.MarkdownRenderer;

/// <summary>
/// Post-processes the HTML produced by <see cref="MarkdownRenderer.RenderBodyHtml"/>. This is intended as an extension
/// point for additional diagram/chart renderers or content normalization in host apps.
/// </summary>
public delegate string MarkdownHtmlPostProcessor(string html, MarkdownRendererOptions options);

