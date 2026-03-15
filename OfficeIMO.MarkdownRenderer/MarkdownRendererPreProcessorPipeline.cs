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
    /// <returns>Processed markdown.</returns>
    public static string Apply(string? markdown, MarkdownRendererOptions options) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        var value = markdown ?? string.Empty;
        if (value.Length == 0) {
            return value;
        }

        var processors = options.MarkdownPreProcessors;
        for (var i = 0; i < processors.Count; i++) {
            var processor = processors[i];
            if (processor == null) {
                continue;
            }

            value = processor(value, options) ?? value;
        }

        return value;
    }
}
