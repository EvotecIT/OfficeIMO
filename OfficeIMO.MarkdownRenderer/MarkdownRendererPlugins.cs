using OfficeIMO.Markdown;

namespace OfficeIMO.MarkdownRenderer;

/// <summary>
/// Built-in renderer plugins. These keep host-specific fence packs layered above the generic
/// <see cref="OfficeIMO.Markdown"/> AST and renderer surface.
/// </summary>
public static class MarkdownRendererPlugins {
    /// <summary>
    /// Generic visual fenced-block plugin for <c>chart</c>, <c>network</c>/<c>visnetwork</c>, and <c>dataview</c>.
    /// This is the default pack applied by <see cref="MarkdownRendererOptions"/>.
    /// </summary>
    public static MarkdownRendererPlugin GenericVisuals { get; } = new MarkdownRendererPlugin(
        "Generic Visuals",
        new Func<MarkdownFencedCodeBlockRenderer>[] {
            MarkdownRendererBuiltInFencedCodeBlocks.CreateChartRenderer,
            MarkdownRendererBuiltInFencedCodeBlocks.CreateNetworkRenderer,
            MarkdownRendererBuiltInFencedCodeBlocks.CreateGenericDataViewRenderer
        });

    /// <summary>
    /// IntelligenceX visual alias plugin that layers <c>ix-chart</c>, <c>ix-network</c>, and <c>ix-dataview</c>
    /// on top of the generic visual pack.
    /// </summary>
    public static MarkdownRendererPlugin IntelligenceXVisuals { get; } = new MarkdownRendererPlugin(
        "IntelligenceX Visuals",
        new Func<MarkdownFencedCodeBlockRenderer>[] {
            MarkdownRendererBuiltInFencedCodeBlocks.CreateIxChartRenderer,
            MarkdownRendererBuiltInFencedCodeBlocks.CreateIxNetworkRenderer,
            MarkdownRendererBuiltInFencedCodeBlocks.CreateDataViewRenderer
        });

    /// <summary>
    /// IntelligenceX transcript plugin that layers the IX transcript reader/AST contract on top of IX visual aliases.
    /// This keeps the transcript-specific parser behavior reusable outside the higher-level feature pack.
    /// </summary>
    public static MarkdownRendererPlugin IntelligenceXTranscriptVisuals { get; } = new MarkdownRendererPlugin(
        "IntelligenceX Transcript Visuals",
        new[] { IntelligenceXVisuals },
        apply: options => MarkdownTranscriptPreparation.ApplyIntelligenceXTranscriptReaderContract(
            options.ReaderOptions,
            preservesGroupedDefinitionLikeParagraphs: false,
            visualFenceLanguageMode: MarkdownVisualFenceLanguageMode.IntelligenceXAliasFence));
}
