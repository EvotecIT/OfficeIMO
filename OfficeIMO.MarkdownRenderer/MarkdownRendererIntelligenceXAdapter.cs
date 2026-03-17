using OfficeIMO.Markdown;

namespace OfficeIMO.MarkdownRenderer;

/// <summary>
/// Opt-in adapter that adds IntelligenceX-specific fenced block aliases on top of the generic renderer surface.
/// </summary>
public static class MarkdownRendererIntelligenceXAdapter {
    /// <summary>
    /// First-class IntelligenceX visual plugin layered on top of the generic OfficeIMO markdown renderer surface.
    /// </summary>
    public static MarkdownRendererPlugin Plugin => MarkdownRendererPlugins.IntelligenceXVisuals;

    /// <summary>
    /// Registers the IntelligenceX alias fences (<c>ix-chart</c>, <c>ix-network</c>, <c>ix-dataview</c>)
    /// if they are not already present.
    /// </summary>
    public static void Apply(MarkdownRendererOptions options) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        options.ApplyPlugin(Plugin);
    }

    /// <summary>
    /// Returns <see langword="true"/> when any IntelligenceX alias fence registration is present.
    /// </summary>
    public static bool IsApplied(MarkdownRendererOptions options) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        return options.HasPlugin(Plugin);
    }
}
