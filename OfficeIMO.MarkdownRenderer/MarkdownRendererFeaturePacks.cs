using OfficeIMO.Markdown;

namespace OfficeIMO.MarkdownRenderer;

/// <summary>
/// Built-in host-level feature packs that coordinate plugins and compatibility behavior.
/// </summary>
public static class MarkdownRendererFeaturePacks {
    /// <summary>
    /// IntelligenceX transcript compatibility contract composed from IX visual aliases and
    /// legacy transcript migration helpers.
    /// </summary>
    public static MarkdownRendererFeaturePack IntelligenceXTranscriptCompatibility { get; } = new MarkdownRendererFeaturePack(
        "officeimo.intelligencex.transcript-compatibility",
        "IntelligenceX Transcript Compatibility",
        new[] { MarkdownRendererPlugins.IntelligenceXTranscriptVisuals },
        options => MarkdownRendererIntelligenceXLegacyMigration.Apply(options));
}
