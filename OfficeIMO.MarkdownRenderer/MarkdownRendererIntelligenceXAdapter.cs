using OfficeIMO.Markdown;

namespace OfficeIMO.MarkdownRenderer;

/// <summary>
/// Opt-in adapter that adds IntelligenceX-specific fenced block aliases on top of the generic renderer surface.
/// </summary>
public static class MarkdownRendererIntelligenceXAdapter {
    /// <summary>
    /// Registers the IntelligenceX alias fences (<c>ix-chart</c>, <c>ix-network</c>, <c>ix-dataview</c>)
    /// if they are not already present.
    /// </summary>
    public static void Apply(MarkdownRendererOptions options) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        AddIfMissing(options, "ix-chart", MarkdownRendererBuiltInFencedCodeBlocks.CreateIxChartRenderer);
        AddIfMissing(options, "ix-network", MarkdownRendererBuiltInFencedCodeBlocks.CreateIxNetworkRenderer);
        AddIfMissing(options, "ix-dataview", MarkdownRendererBuiltInFencedCodeBlocks.CreateDataViewRenderer);
    }

    /// <summary>
    /// Returns <see langword="true"/> when any IntelligenceX alias fence registration is present.
    /// </summary>
    public static bool IsApplied(MarkdownRendererOptions options) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        return HasLanguage(options, "ix-chart")
            || HasLanguage(options, "ix-network")
            || HasLanguage(options, "ix-dataview");
    }

    private static void AddIfMissing(
        MarkdownRendererOptions options,
        string language,
        Func<MarkdownFencedCodeBlockRenderer> factory) {
        if (!HasLanguage(options, language)) {
            options.FencedCodeBlockRenderers.Add(factory());
        }
    }

    private static bool HasLanguage(MarkdownRendererOptions options, string language) {
        var renderers = options.FencedCodeBlockRenderers;
        for (int i = 0; i < renderers.Count; i++) {
            var renderer = renderers[i];
            if (renderer == null) {
                continue;
            }

            var languages = renderer.Languages;
            for (int j = 0; j < languages.Count; j++) {
                if (string.Equals(languages[j], language, StringComparison.OrdinalIgnoreCase)) {
                    return true;
                }
            }
        }

        return false;
    }
}
