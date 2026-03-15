using OfficeIMO.Markdown;

namespace OfficeIMO.MarkdownRenderer;

/// <summary>
/// Convenience factory methods and composition helpers for common markdown host scenarios.
/// These are intentionally opinionated, but still fully configurable via <see cref="MarkdownRendererOptions"/>.
/// </summary>
public static class MarkdownRendererPresets {
    private static MarkdownReaderOptions CreateStrictReaderOptions(MarkdownReaderOptions.MarkdownDialectProfile readerProfile) {
        var reader = MarkdownReaderOptions.CreateProfile(readerProfile);
        ApplyStrictReaderSecurityDefaults(reader);
        return reader;
    }

    private static void ApplyStrictReaderSecurityDefaults(MarkdownReaderOptions reader) {
        reader.HtmlBlocks = false;
        // Keep block HTML disabled, but allow the reader's safe inline-tag subset
        // so built-in formatting like <u>...</u> and <br> can survive strict rendering.
        reader.InlineHtml = true;
        reader.DisallowFileUrls = true;
        reader.AllowDataUrls = false;
        reader.AllowProtocolRelativeUrls = false;
        reader.RestrictUrlSchemes = true;
        reader.AllowedUrlSchemes = new[] { "http", "https", "mailto" };
    }

    private static void ApplyStrictRenderingDefaults(
        MarkdownRendererOptions options,
        string? baseHref,
        MarkdownReaderOptions.MarkdownDialectProfile readerProfile) {
        options.BaseHref = baseHref;
        options.ReaderOptions = CreateStrictReaderOptions(readerProfile);

        options.NormalizeSoftWrappedStrongSpans = true;
        options.NormalizeInlineCodeSpanLineBreaks = true;
        options.NormalizeEscapedInlineCodeSpans = true;
        options.NormalizeTightStrongBoundaries = true;
        options.NormalizeTightArrowStrongBoundaries = true;
        options.NormalizeBrokenStrongArrowLabels = true;
        options.NormalizeSignalFlowLabelSpacing = true;
        options.NormalizeTightColonSpacing = true;
        options.NormalizeHeadingListBoundaries = true;
        options.NormalizeCompactStrongLabelListBoundaries = true;
        options.NormalizeCompactHeadingBoundaries = true;
        options.NormalizeColonListBoundaries = true;
        options.NormalizeCompactFenceBodyBoundaries = true;
        options.NormalizeLooseStrongDelimiters = true;
        options.NormalizeOrderedListMarkerSpacing = true;
        options.NormalizeOrderedListParenMarkers = true;
        options.NormalizeOrderedListCaretArtifacts = true;
        options.NormalizeTightParentheticalSpacing = true;
        options.NormalizeNestedStrongDelimiters = true;

        options.HtmlOptions.RawHtmlHandling = RawHtmlHandling.Strip;
        options.HtmlOptions.ExternalLinksTargetBlank = true;
        options.HtmlOptions.ExternalLinksRel = "noopener noreferrer nofollow ugc";
        options.HtmlOptions.ExternalLinksReferrerPolicy = "no-referrer";

        options.HtmlOptions.RestrictHttpLinksToBaseOrigin = true;
        options.HtmlOptions.RestrictHttpImagesToBaseOrigin = true;
        options.HtmlOptions.BlockExternalHttpImages = true;

        options.HtmlOptions.ImagesLoadingLazy = true;
        options.HtmlOptions.ImagesDecodingAsync = true;
        options.HtmlOptions.ImagesReferrerPolicy = "no-referrer";

        options.MaxMarkdownChars = 500_000;
        options.MaxBodyHtmlBytes = 5_000_000;
        options.MarkdownOverflowHandling = OverflowHandling.Truncate;
        options.BodyHtmlOverflowHandling = OverflowHandling.RenderError;
    }

    private static void ApplyMinimalRenderingDefaults(MarkdownRendererOptions options) {
        options.EnableCodeCopyButtons = false;
        options.EnableTableCopyButtons = false;

        options.Mermaid.Enabled = false;
        options.Chart.Enabled = false;
        options.Network.Enabled = false;
        options.Math.Enabled = false;
        if (options.HtmlOptions.Prism != null) {
            options.HtmlOptions.Prism.Enabled = false;
        }
    }

    private static void ApplyIntelligenceXTranscriptNormalizationDefaults(MarkdownRendererOptions options) {
        options.NormalizeZeroWidthSpacingArtifacts = true;
        options.NormalizeEmojiWordJoins = true;
        options.NormalizeCompactNumberedChoiceBoundaries = true;
        options.NormalizeSentenceCollapsedBullets = true;
        options.NormalizeWrappedSignalFlowStrongRuns = true;
        options.NormalizeSignalFlowLabelSpacing = true;
        options.NormalizeCollapsedMetricChains = true;
        options.NormalizeHostLabelBulletArtifacts = true;
        options.NormalizeCollapsedOrderedListBoundaries = true;
        options.NormalizeOrderedListStrongDetailClosures = true;
        options.NormalizeStandaloneHashHeadingSeparators = true;
        options.NormalizeBrokenTwoLineStrongLeadIns = true;
        options.NormalizeDanglingTrailingStrongListClosers = true;
        options.NormalizeMetricValueStrongRuns = true;
    }

    private static void ApplyIntelligenceXTranscriptDocumentTransforms(MarkdownRendererOptions options) {
        var transforms = options.ReaderOptions.DocumentTransforms;
        bool hasDefinitionCompatibility = false;
        bool hasVisualUpgrade = false;

        for (var i = 0; i < transforms.Count; i++) {
            switch (transforms[i]) {
                case MarkdownSimpleDefinitionListParagraphTransform:
                    hasDefinitionCompatibility = true;
                    break;
                case MarkdownJsonVisualCodeBlockTransform existing
                    when existing.LanguageMode == MarkdownVisualFenceLanguageMode.IntelligenceXAliasFence:
                    hasVisualUpgrade = true;
                    break;
            }
        }

        if (!hasDefinitionCompatibility) {
            transforms.Add(new MarkdownSimpleDefinitionListParagraphTransform());
        }

        if (!hasVisualUpgrade) {
            transforms.Add(new MarkdownJsonVisualCodeBlockTransform(MarkdownVisualFenceLanguageMode.IntelligenceXAliasFence));
        }
    }

    private static void ApplyIntelligenceXTranscriptReaderContract(
        MarkdownRendererOptions options,
        MarkdownReaderOptions.MarkdownDialectProfile readerProfile) {
        var transcriptReader = MarkdownTranscriptPreparation.CreateIntelligenceXTranscriptReaderOptions(
            readerProfile,
            preservesGroupedDefinitionLikeParagraphs: false);

        ApplyStrictReaderSecurityDefaults(transcriptReader);
        options.ReaderOptions = transcriptReader;
        ApplyIntelligenceXTranscriptDocumentTransforms(options);
    }

    /// <summary>
    /// Applies portable HTML output fallbacks so OfficeIMO-only block chrome degrades to simpler generic HTML.
    /// </summary>
    public static void ApplyPortableHtmlOutputProfile(MarkdownRendererOptions options) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        MarkdownBlockRenderBuiltInExtensions.AddPortableHtmlFallbacks(options.HtmlOptions);
    }

    /// <summary>
    /// Applies chat-oriented presentation defaults on top of an existing renderer preset.
    /// This only affects host presentation/chrome and copy-button UX; it does not change the security profile.
    /// </summary>
    public static void ApplyChatPresentation(MarkdownRendererOptions options, bool enableCopyButtons = true) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        options.HtmlOptions.Style = HtmlStyle.ChatAuto;
        options.HtmlOptions.CssScopeSelector = "#omdRoot article.markdown-body";
        options.EnableCodeCopyButtons = enableCopyButtons;
        options.EnableTableCopyButtons = enableCopyButtons;
    }

    /// <summary>
    /// Strict preset for generic untrusted markdown hosting.
    /// - Disables HTML parsing (blocks + inline)
    /// - Strips any raw HTML blocks
    /// - Restricts URL schemes and blocks file/data/protocol-relative URLs
    /// - Blocks external HTTP(S) images unless same-origin with BaseHref/BaseUri
    /// </summary>
    public static MarkdownRendererOptions CreateStrict(string? baseHref = null) {
        return CreateStrict(MarkdownReaderOptions.MarkdownDialectProfile.OfficeIMO, baseHref);
    }

    /// <summary>
    /// Strict preset for generic untrusted markdown hosting using an explicit reader profile.
    /// This keeps the same security defaults while allowing hosts to target OfficeIMO, CommonMark,
    /// GitHub Flavored Markdown, or the portable OfficeIMO subset.
    /// </summary>
    public static MarkdownRendererOptions CreateStrict(MarkdownReaderOptions.MarkdownDialectProfile readerProfile, string? baseHref = null) {
        var options = new MarkdownRendererOptions();
        ApplyStrictRenderingDefaults(options, baseHref, readerProfile);
        if (readerProfile == MarkdownReaderOptions.MarkdownDialectProfile.Portable) {
            ApplyPortableHtmlOutputProfile(options);
        }
        return options;
    }

    /// <summary>
    /// Strict preset for generic untrusted markdown hosting with the portable reader profile enabled.
    /// This disables OfficeIMO-only literal autolinks, callouts, and task-list parsing while keeping the same security defaults.
    /// </summary>
    public static MarkdownRendererOptions CreateStrictPortable(string? baseHref = null) =>
        CreateStrict(MarkdownReaderOptions.MarkdownDialectProfile.Portable, baseHref);

    /// <summary>
    /// Strict preset for generic untrusted markdown hosting with optional client-side renderers disabled.
    /// This disables Mermaid, charts, math, Prism, and copy-button helpers to minimize shell scripting.
    /// </summary>
    public static MarkdownRendererOptions CreateStrictMinimal(string? baseHref = null) {
        return CreateStrictMinimal(MarkdownReaderOptions.MarkdownDialectProfile.OfficeIMO, baseHref);
    }

    /// <summary>
    /// Strict minimal preset for generic untrusted markdown hosting using an explicit reader profile.
    /// </summary>
    public static MarkdownRendererOptions CreateStrictMinimal(MarkdownReaderOptions.MarkdownDialectProfile readerProfile, string? baseHref = null) {
        var options = CreateStrict(readerProfile, baseHref);
        ApplyMinimalRenderingDefaults(options);
        return options;
    }

    /// <summary>
    /// Strict minimal preset for generic untrusted markdown hosting with the portable reader profile enabled.
    /// </summary>
    public static MarkdownRendererOptions CreateStrictMinimalPortable(string? baseHref = null) =>
        CreateStrictMinimal(MarkdownReaderOptions.MarkdownDialectProfile.Portable, baseHref);

    /// <summary>
    /// Relaxed preset for trusted or controlled generic markdown hosting.
    /// - Allows HTML parsing but sanitizes raw HTML blocks
    /// - Allows external HTTP(S) images unless further restricted by the host
    /// </summary>
    public static MarkdownRendererOptions CreateRelaxed(string? baseHref = null) {
        return CreateRelaxed(MarkdownReaderOptions.MarkdownDialectProfile.OfficeIMO, baseHref);
    }

    /// <summary>
    /// Relaxed preset for trusted or controlled generic markdown hosting using an explicit reader profile.
    /// </summary>
    public static MarkdownRendererOptions CreateRelaxed(MarkdownReaderOptions.MarkdownDialectProfile readerProfile, string? baseHref = null) {
        var options = CreateStrict(readerProfile, baseHref);

        options.ReaderOptions.HtmlBlocks = true;
        options.ReaderOptions.InlineHtml = true;

        options.HtmlOptions.RawHtmlHandling = RawHtmlHandling.Sanitize;
        options.HtmlOptions.BlockExternalHttpImages = false;
        options.HtmlOptions.RestrictHttpLinksToBaseOrigin = false;
        options.HtmlOptions.RestrictHttpImagesToBaseOrigin = false;

        return options;
    }

    /// <summary>
    /// Strict preset for the explicit IntelligenceX transcript rendering contract.
    /// This is the preferred first-class name for the OfficeIMO-hosted IX transcript surface.
    /// </summary>
    public static MarkdownRendererOptions CreateIntelligenceXTranscript(string? baseHref = null) {
        return CreateIntelligenceXTranscript(MarkdownReaderOptions.MarkdownDialectProfile.OfficeIMO, baseHref);
    }

    /// <summary>
    /// Strict preset for the explicit IntelligenceX transcript rendering contract using an explicit reader profile.
    /// </summary>
    public static MarkdownRendererOptions CreateIntelligenceXTranscript(MarkdownReaderOptions.MarkdownDialectProfile readerProfile, string? baseHref = null) {
        var options = CreateStrict(readerProfile, baseHref);
        ApplyIntelligenceXTranscriptNormalizationDefaults(options);
        ApplyIntelligenceXTranscriptReaderContract(options, readerProfile);
        ApplyChatPresentation(options, enableCopyButtons: true);
        MarkdownRendererIntelligenceXAdapter.Apply(options);
        MarkdownRendererIntelligenceXLegacyMigration.Apply(options);
        return options;
    }

    /// <summary>
    /// Strict preset for the explicit IntelligenceX transcript contract with the portable reader profile enabled.
    /// This disables OfficeIMO-only literal autolinks, callouts, and task-list parsing while keeping the same transcript security defaults.
    /// </summary>
    public static MarkdownRendererOptions CreateIntelligenceXTranscriptPortable(string? baseHref = null) =>
        CreateIntelligenceXTranscript(MarkdownReaderOptions.MarkdownDialectProfile.Portable, baseHref);

    /// <summary>
    /// Strict minimal preset for the explicit IntelligenceX transcript rendering contract.
    /// This keeps the IX transcript semantics while minimizing client-side shell features by default.
    /// </summary>
    public static MarkdownRendererOptions CreateIntelligenceXTranscriptMinimal(string? baseHref = null) {
        return CreateIntelligenceXTranscriptMinimal(MarkdownReaderOptions.MarkdownDialectProfile.OfficeIMO, baseHref);
    }

    /// <summary>
    /// Strict minimal preset for the explicit IntelligenceX transcript rendering contract using an explicit reader profile.
    /// </summary>
    public static MarkdownRendererOptions CreateIntelligenceXTranscriptMinimal(MarkdownReaderOptions.MarkdownDialectProfile readerProfile, string? baseHref = null) {
        var options = CreateStrictMinimal(readerProfile, baseHref);
        ApplyIntelligenceXTranscriptNormalizationDefaults(options);
        ApplyIntelligenceXTranscriptDocumentTransforms(options);
        ApplyChatPresentation(options, enableCopyButtons: false);
        MarkdownRendererIntelligenceXAdapter.Apply(options);
        MarkdownRendererIntelligenceXLegacyMigration.Apply(options);
        return options;
    }

    /// <summary>
    /// Strict minimal preset for the explicit IntelligenceX transcript contract with the portable reader profile enabled.
    /// This combines the minimal shell-friendly renderer defaults with the stricter reader preset used for portability-sensitive hosts.
    /// </summary>
    public static MarkdownRendererOptions CreateIntelligenceXTranscriptMinimalPortable(string? baseHref = null) =>
        CreateIntelligenceXTranscriptMinimal(MarkdownReaderOptions.MarkdownDialectProfile.Portable, baseHref);

    /// <summary>
    /// Strict desktop-shell preset for the explicit IntelligenceX transcript contract.
    /// This keeps the minimal chat-shell defaults while enabling the interactive visual surface
    /// required by the IntelligenceX desktop host.
    /// </summary>
    public static MarkdownRendererOptions CreateIntelligenceXTranscriptDesktopShell(string? baseHref = null) {
        return CreateIntelligenceXTranscriptDesktopShell(MarkdownReaderOptions.MarkdownDialectProfile.OfficeIMO, baseHref);
    }

    /// <summary>
    /// Strict desktop-shell preset for the explicit IntelligenceX transcript contract using an explicit reader profile.
    /// </summary>
    public static MarkdownRendererOptions CreateIntelligenceXTranscriptDesktopShell(MarkdownReaderOptions.MarkdownDialectProfile readerProfile, string? baseHref = null) {
        var options = CreateIntelligenceXTranscriptMinimal(readerProfile, baseHref);
        options.Mermaid.Enabled = true;
        options.Chart.Enabled = true;
        options.Network.Enabled = true;
        return options;
    }

    /// <summary>
    /// Relaxed preset for trusted or controlled transcript content rendered in a WebView.
    /// - Allows HTML parsing but sanitizes raw HTML blocks (very conservative allowlist)
    /// - Allows external HTTP(S) images (unless further restricted via host/origin allowlists)
    /// </summary>
    public static MarkdownRendererOptions CreateIntelligenceXTranscriptRelaxed(string? baseHref = null) {
        return CreateIntelligenceXTranscriptRelaxed(MarkdownReaderOptions.MarkdownDialectProfile.OfficeIMO, baseHref);
    }

    /// <summary>
    /// Relaxed preset for trusted or controlled IntelligenceX transcript content using an explicit reader profile.
    /// </summary>
    public static MarkdownRendererOptions CreateIntelligenceXTranscriptRelaxed(MarkdownReaderOptions.MarkdownDialectProfile readerProfile, string? baseHref = null) {
        var options = CreateRelaxed(readerProfile, baseHref);
        ApplyIntelligenceXTranscriptReaderContract(options, readerProfile);
        options.ReaderOptions.HtmlBlocks = true;
        options.ReaderOptions.InlineHtml = true;
        ApplyChatPresentation(options, enableCopyButtons: true);
        MarkdownRendererIntelligenceXAdapter.Apply(options);
        MarkdownRendererIntelligenceXLegacyMigration.Apply(options);
        return options;
    }
}
