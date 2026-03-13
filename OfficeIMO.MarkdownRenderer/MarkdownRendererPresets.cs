using OfficeIMO.Markdown;

namespace OfficeIMO.MarkdownRenderer;

/// <summary>
/// Convenience factory methods and composition helpers for common markdown host scenarios.
/// These are intentionally opinionated, but still fully configurable via <see cref="MarkdownRendererOptions"/>.
/// </summary>
public static class MarkdownRendererPresets {
    private static MarkdownReaderOptions CreateStrictReaderOptions(bool portableProfile) {
        var reader = portableProfile
            ? MarkdownReaderOptions.CreatePortableProfile()
            : new MarkdownReaderOptions();

        reader.HtmlBlocks = false;
        reader.InlineHtml = false;
        reader.DisallowFileUrls = true;
        reader.AllowDataUrls = false;
        reader.AllowProtocolRelativeUrls = false;
        reader.RestrictUrlSchemes = true;
        reader.AllowedUrlSchemes = new[] { "http", "https", "mailto" };

        return reader;
    }

    private static void ApplyStrictRenderingDefaults(MarkdownRendererOptions options, string? baseHref, bool portableProfile) {
        options.BaseHref = baseHref;
        options.ReaderOptions = CreateStrictReaderOptions(portableProfile);

        options.NormalizeSoftWrappedStrongSpans = true;
        options.NormalizeInlineCodeSpanLineBreaks = true;
        options.NormalizeEscapedInlineCodeSpans = true;
        options.NormalizeTightStrongBoundaries = true;
        options.NormalizeTightArrowStrongBoundaries = true;
        options.NormalizeBrokenStrongArrowLabels = true;
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
        var options = new MarkdownRendererOptions();
        ApplyStrictRenderingDefaults(options, baseHref, portableProfile: false);
        return options;
    }

    /// <summary>
    /// Strict preset for generic untrusted markdown hosting with the portable reader profile enabled.
    /// This disables OfficeIMO-only literal autolinks, callouts, and task-list parsing while keeping the same security defaults.
    /// </summary>
    public static MarkdownRendererOptions CreateStrictPortable(string? baseHref = null) {
        var options = new MarkdownRendererOptions();
        ApplyStrictRenderingDefaults(options, baseHref, portableProfile: true);
        return options;
    }

    /// <summary>
    /// Strict preset for generic untrusted markdown hosting with optional client-side renderers disabled.
    /// This disables Mermaid, charts, math, Prism, and copy-button helpers to minimize shell scripting.
    /// </summary>
    public static MarkdownRendererOptions CreateStrictMinimal(string? baseHref = null) {
        var options = CreateStrict(baseHref);
        ApplyMinimalRenderingDefaults(options);
        return options;
    }

    /// <summary>
    /// Strict minimal preset for generic untrusted markdown hosting with the portable reader profile enabled.
    /// </summary>
    public static MarkdownRendererOptions CreateStrictMinimalPortable(string? baseHref = null) {
        var options = CreateStrictPortable(baseHref);
        ApplyMinimalRenderingDefaults(options);
        return options;
    }

    /// <summary>
    /// Relaxed preset for trusted or controlled generic markdown hosting.
    /// - Allows HTML parsing but sanitizes raw HTML blocks
    /// - Allows external HTTP(S) images unless further restricted by the host
    /// </summary>
    public static MarkdownRendererOptions CreateRelaxed(string? baseHref = null) {
        var options = CreateStrict(baseHref);

        options.ReaderOptions.HtmlBlocks = true;
        options.ReaderOptions.InlineHtml = true;

        options.HtmlOptions.RawHtmlHandling = RawHtmlHandling.Sanitize;
        options.HtmlOptions.BlockExternalHttpImages = false;
        options.HtmlOptions.RestrictHttpLinksToBaseOrigin = false;
        options.HtmlOptions.RestrictHttpImagesToBaseOrigin = false;

        return options;
    }

    /// <summary>
    /// Strict preset for untrusted chat messages.
    /// - Disables HTML parsing (blocks + inline)
    /// - Strips any raw HTML blocks
    /// - Restricts URL schemes and blocks file/data/protocol-relative URLs
    /// - Blocks external HTTP(S) images unless same-origin with BaseHref/BaseUri
    /// </summary>
    public static MarkdownRendererOptions CreateChatStrict(string? baseHref = null) {
        var options = CreateStrict(baseHref);
        ApplyChatPresentation(options, enableCopyButtons: true);
        MarkdownRendererIntelligenceXAdapter.Apply(options);
        return options;
    }

    /// <summary>
    /// Strict preset for untrusted chat messages, but with the portable reader profile enabled.
    /// This disables OfficeIMO-only literal autolinks, callouts, and task-list parsing while keeping the same chat security defaults.
    /// </summary>
    public static MarkdownRendererOptions CreateChatStrictPortable(string? baseHref = null) {
        var options = CreateStrictPortable(baseHref);
        ApplyChatPresentation(options, enableCopyButtons: true);
        MarkdownRendererIntelligenceXAdapter.Apply(options);
        return options;
    }

    /// <summary>
    /// Strict preset for untrusted chat messages, with optional client-side renderers disabled.
    /// This disables Mermaid/Chart/Math/Prism and the copy-button UX helpers to minimize script usage in the shell.
    /// </summary>
    public static MarkdownRendererOptions CreateChatStrictMinimal(string? baseHref = null) {
        var options = CreateStrictMinimal(baseHref);
        ApplyChatPresentation(options, enableCopyButtons: false);
        MarkdownRendererIntelligenceXAdapter.Apply(options);
        return options;
    }

    /// <summary>
    /// Strict minimal preset for untrusted chat messages, with the portable reader profile enabled.
    /// This combines the minimal shell-friendly renderer defaults with the stricter reader preset used for portability-sensitive hosts.
    /// </summary>
    public static MarkdownRendererOptions CreateChatStrictMinimalPortable(string? baseHref = null) {
        var options = CreateStrictMinimalPortable(baseHref);
        ApplyChatPresentation(options, enableCopyButtons: false);
        MarkdownRendererIntelligenceXAdapter.Apply(options);
        return options;
    }

    /// <summary>
    /// Relaxed preset for trusted/controlled content rendered in a WebView.
    /// - Allows HTML parsing but sanitizes raw HTML blocks (very conservative allowlist)
    /// - Allows external HTTP(S) images (unless further restricted via host/origin allowlists)
    /// </summary>
    public static MarkdownRendererOptions CreateChatRelaxed(string? baseHref = null) {
        var options = CreateRelaxed(baseHref);
        ApplyChatPresentation(options, enableCopyButtons: true);
        MarkdownRendererIntelligenceXAdapter.Apply(options);
        return options;
    }
}
