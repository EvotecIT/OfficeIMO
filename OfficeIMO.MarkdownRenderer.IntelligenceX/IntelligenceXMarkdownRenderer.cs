using OfficeIMO.Markdown;
using OfficeIMO.MarkdownRenderer;

namespace OfficeIMO.MarkdownRenderer.IntelligenceX;

/// <summary>
/// First-party IntelligenceX plugin entrypoint layered on top of <see cref="OfficeIMO.MarkdownRenderer"/>.
/// </summary>
public static class IntelligenceXMarkdownRenderer {
    /// <summary>
    /// IntelligenceX visual alias plugin that adds <c>ix-chart</c>, <c>ix-network</c>, and <c>ix-dataview</c>.
    /// </summary>
    public static MarkdownRendererPlugin VisualsPlugin { get; } = new MarkdownRendererPlugin(
        "IntelligenceX Visuals",
        new[] { MarkdownRendererPlugins.IntelligenceXVisuals },
        new[] { IntelligenceXVisualFenceSchemas.Visuals });

    /// <summary>
    /// IntelligenceX transcript plugin that carries the IX visual aliases, fence-option schema,
    /// and transcript reader/AST contract without the legacy transcript cleanup preprocessors.
    /// </summary>
    public static MarkdownRendererPlugin TranscriptPlugin { get; } = new MarkdownRendererPlugin(
        "IntelligenceX Transcript",
        new[] { VisualsPlugin },
        apply: options => MarkdownTranscriptPreparation.ApplyIntelligenceXTranscriptReaderContract(
            options.ReaderOptions,
            preservesGroupedDefinitionLikeParagraphs: false,
            visualFenceLanguageMode: MarkdownVisualFenceLanguageMode.IntelligenceXAliasFence));

    /// <summary>
    /// IntelligenceX transcript compatibility feature pack composed from IX visual aliases and
    /// legacy transcript migration helpers.
    /// </summary>
    public static MarkdownRendererFeaturePack TranscriptCompatibilityPack { get; } = new MarkdownRendererFeaturePack(
        "officeimo.intelligencex.transcript-compatibility",
        "IntelligenceX Transcript Compatibility",
        new[] { TranscriptPlugin },
        options => MarkdownRendererIntelligenceXLegacyMigration.Apply(options));

    /// <summary>
    /// IntelligenceX visual fence option schema layered on top of the shared renderer contract.
    /// </summary>
    public static MarkdownFenceOptionSchema VisualFenceSchema => IntelligenceXVisualFenceSchemas.Visuals;

    /// <summary>
    /// Applies the IntelligenceX visual alias plugin.
    /// </summary>
    public static void ApplyVisuals(MarkdownRendererOptions options) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        options.ApplyPlugin(VisualsPlugin);
    }

    /// <summary>
    /// Applies the IntelligenceX transcript reader/AST contract together with IX visual aliases and schema support.
    /// </summary>
    public static void ApplyTranscriptContract(MarkdownRendererOptions options) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        options.ApplyPlugin(TranscriptPlugin);
    }

    /// <summary>
    /// Applies the IntelligenceX visual aliases and transcript legacy migration helpers.
    /// </summary>
    public static void ApplyTranscriptCompatibility(MarkdownRendererOptions options) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        options.ApplyFeaturePack(TranscriptCompatibilityPack);
    }

    /// <summary>
    /// Returns <see langword="true"/> when the IntelligenceX transcript plugin is already applied.
    /// </summary>
    public static bool HasTranscriptContract(MarkdownRendererOptions options) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        return options.HasPlugin(TranscriptPlugin);
    }

    /// <summary>
    /// Returns <see langword="true"/> when the IntelligenceX transcript compatibility pack is already applied.
    /// </summary>
    public static bool HasTranscriptCompatibility(MarkdownRendererOptions options) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        return options.HasFeaturePack(TranscriptCompatibilityPack);
    }

    /// <summary>
    /// Applies the IntelligenceX visual fence option schema.
    /// </summary>
    public static void ApplyVisualFenceSchema(MarkdownRendererOptions options) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        options.ApplyFenceOptionSchema(VisualFenceSchema);
    }

    /// <summary>
    /// Returns <see langword="true"/> when the IntelligenceX visual fence option schema is already registered.
    /// </summary>
    public static bool HasVisualFenceSchema(MarkdownRendererOptions options) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        return options.HasFenceOptionSchema(VisualFenceSchema);
    }

    /// <summary>
    /// Parses typed IntelligenceX visual fence metadata from a raw fenced-code info string.
    /// </summary>
    public static IntelligenceXVisualFenceOptions ParseVisualFenceOptions(string? infoString) =>
        IntelligenceXVisualFenceOptions.Parse(infoString);

    /// <summary>
    /// Parses typed IntelligenceX visual fence metadata from a shared fenced-code info descriptor.
    /// </summary>
    public static IntelligenceXVisualFenceOptions ParseVisualFenceOptions(MarkdownCodeFenceInfo? fenceInfo) =>
        IntelligenceXVisualFenceOptions.Parse(fenceInfo);

    /// <summary>
    /// Creates the strict IntelligenceX transcript preset.
    /// </summary>
    public static MarkdownRendererOptions CreateTranscript(string? baseHref = null) {
        var options = MarkdownRendererPresets.CreateIntelligenceXTranscript(baseHref);
        ApplyVisuals(options);
        return options;
    }

    /// <summary>
    /// Creates the strict minimal IntelligenceX transcript preset.
    /// </summary>
    public static MarkdownRendererOptions CreateTranscriptMinimal(string? baseHref = null) {
        var options = MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimal(baseHref);
        ApplyVisuals(options);
        return options;
    }

    /// <summary>
    /// Creates the strict desktop-shell IntelligenceX transcript preset.
    /// </summary>
    public static MarkdownRendererOptions CreateTranscriptDesktopShell(string? baseHref = null) {
        var options = MarkdownRendererPresets.CreateIntelligenceXTranscriptDesktopShell(baseHref);
        ApplyVisuals(options);
        return options;
    }

    /// <summary>
    /// Creates the relaxed IntelligenceX transcript preset.
    /// </summary>
    public static MarkdownRendererOptions CreateTranscriptRelaxed(string? baseHref = null) {
        var options = MarkdownRendererPresets.CreateIntelligenceXTranscriptRelaxed(baseHref);
        ApplyVisuals(options);
        return options;
    }
}
