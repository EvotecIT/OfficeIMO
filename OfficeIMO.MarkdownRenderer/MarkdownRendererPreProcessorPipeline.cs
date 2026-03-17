namespace OfficeIMO.MarkdownRenderer;

/// <summary>
/// Helpers for applying renderer markdown pre-processors outside the full HTML render path.
/// </summary>
public static class MarkdownRendererPreProcessorPipeline {
    /// <summary>
    /// Applies the configured markdown pre-processor chain in order and returns the transformed markdown.
    /// </summary>
    /// <param name="markdown">Markdown input to process.</param>
    /// <param name="options">Renderer options providing the pre-processor chain.</param>
    /// <param name="diagnostics">Optional diagnostics sink describing pre-parse stages that ran.</param>
    /// <returns>Processed markdown.</returns>
    public static string Apply(
        string? markdown,
        MarkdownRendererOptions options,
        ICollection<MarkdownRendererPreProcessorDiagnostic>? diagnostics = null) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        return MarkdownRenderer.ApplyPreParseProcessing(markdown, options, diagnostics);
    }
}
