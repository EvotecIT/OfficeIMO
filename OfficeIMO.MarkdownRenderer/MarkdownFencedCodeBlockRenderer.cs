using OfficeIMO.Markdown;

namespace OfficeIMO.MarkdownRenderer;

/// <summary>
/// Converts a rendered fenced code block into host-specific HTML.
/// Returning <see langword="null"/> preserves the original rendered block.
/// </summary>
public delegate string? MarkdownFencedCodeBlockHtmlRenderer(MarkdownFencedCodeBlockMatch match, MarkdownRendererOptions options);

/// <summary>
/// Builds optional HTML to append into the shell document head for a fenced code block renderer.
/// Returning <see langword="null"/> emits nothing.
/// </summary>
public delegate string? MarkdownRendererShellHeadBuilder(MarkdownRendererOptions options, AssetMode assetMode);

/// <summary>
/// Builds optional JavaScript fragments inserted into the shell update pipeline.
/// Returning <see langword="null"/> emits nothing.
/// </summary>
public delegate string? MarkdownRendererShellUpdateScriptBuilder(MarkdownRendererOptions options);

/// <summary>
/// Defines a custom fenced code block renderer extension.
/// </summary>
public sealed class MarkdownFencedCodeBlockRenderer {
    /// <summary>
    /// Creates a new fenced code block renderer.
    /// </summary>
    public MarkdownFencedCodeBlockRenderer(string name, IEnumerable<string> languages, MarkdownFencedCodeBlockHtmlRenderer renderHtml) {
        if (string.IsNullOrWhiteSpace(name)) {
            throw new ArgumentException("Renderer name is required.", nameof(name));
        }

        if (languages == null) {
            throw new ArgumentNullException(nameof(languages));
        }

        RenderHtml = renderHtml ?? throw new ArgumentNullException(nameof(renderHtml));

        var normalized = new List<string>();
        foreach (var language in languages) {
            var value = (language ?? string.Empty).Trim();
            if (value.Length == 0) {
                continue;
            }

            var exists = false;
            for (int i = 0; i < normalized.Count; i++) {
                if (string.Equals(normalized[i], value, StringComparison.OrdinalIgnoreCase)) {
                    exists = true;
                    break;
                }
            }

            if (!exists) {
                normalized.Add(value);
            }
        }

        if (normalized.Count == 0) {
            throw new ArgumentException("At least one fenced code block language is required.", nameof(languages));
        }

        Name = name.Trim();
        Languages = normalized;
    }

    /// <summary>
    /// Friendly renderer name used for diagnostics and documentation.
    /// </summary>
    public string Name { get; }

    /// <summary>
    /// Fenced code block languages handled by this renderer.
    /// </summary>
    public IReadOnlyList<string> Languages { get; }

    /// <summary>
    /// HTML conversion callback invoked for each matching rendered code block.
    /// </summary>
    public MarkdownFencedCodeBlockHtmlRenderer RenderHtml { get; }

    /// <summary>
    /// Optional shell head HTML builder invoked by <see cref="MarkdownRenderer.BuildShellHtml(string?, MarkdownRendererOptions?)"/>.
    /// </summary>
    public MarkdownRendererShellHeadBuilder? BuildShellHeadHtml { get; set; }

    /// <summary>
    /// Optional JavaScript emitted before <c>root.innerHTML = newBodyHtml;</c> in the shell update pipeline.
    /// </summary>
    public MarkdownRendererShellUpdateScriptBuilder? BuildBeforeContentReplaceScript { get; set; }

    /// <summary>
    /// Optional JavaScript emitted after <c>root.innerHTML = newBodyHtml;</c> in the shell update pipeline.
    /// </summary>
    public MarkdownRendererShellUpdateScriptBuilder? BuildAfterContentReplaceScript { get; set; }
}
