using OfficeIMO.Markdown;

namespace OfficeIMO.MarkdownRenderer;

/// <summary>
/// Convenience factory methods for common WebView/chat scenarios.
/// These are intentionally opinionated, but still fully configurable via <see cref="MarkdownRendererOptions"/>.
/// </summary>
public static class MarkdownRendererPresets {
    /// <summary>
    /// Strict preset for untrusted chat messages.
    /// - Disables HTML parsing (blocks + inline)
    /// - Strips any raw HTML blocks
    /// - Restricts URL schemes and blocks file/data/protocol-relative URLs
    /// - Blocks external HTTP(S) images unless same-origin with BaseHref/BaseUri
    /// </summary>
    public static MarkdownRendererOptions CreateChatStrict(string? baseHref = null) {
        var o = new MarkdownRendererOptions();
        o.BaseHref = baseHref;

        // Chat look-and-feel (compact, embed-friendly). Scope to the renderer root to avoid CSS bleed.
        o.HtmlOptions.Style = HtmlStyle.ChatAuto;
        o.HtmlOptions.CssScopeSelector = "#omdRoot article.markdown-body";

        o.ReaderOptions.HtmlBlocks = false;
        o.ReaderOptions.InlineHtml = false;
        o.ReaderOptions.DisallowFileUrls = true;
        o.ReaderOptions.AllowDataUrls = false;
        o.ReaderOptions.AllowProtocolRelativeUrls = false;
        o.ReaderOptions.RestrictUrlSchemes = true;
        o.ReaderOptions.AllowedUrlSchemes = new[] { "http", "https", "mailto" };
        o.NormalizeSoftWrappedStrongSpans = true;
        o.NormalizeInlineCodeSpanLineBreaks = true;
        o.NormalizeEscapedInlineCodeSpans = true;
        o.NormalizeTightStrongBoundaries = true;
        o.NormalizeLooseStrongDelimiters = true;
        o.NormalizeOrderedListMarkerSpacing = true;
        o.NormalizeOrderedListParenMarkers = true;
        o.NormalizeOrderedListCaretArtifacts = true;
        o.NormalizeTightParentheticalSpacing = true;
        o.NormalizeNestedStrongDelimiters = true;

        o.HtmlOptions.RawHtmlHandling = RawHtmlHandling.Strip;
        o.HtmlOptions.ExternalLinksTargetBlank = true;
        o.HtmlOptions.ExternalLinksRel = "noopener noreferrer nofollow ugc";
        o.HtmlOptions.ExternalLinksReferrerPolicy = "no-referrer";

        o.HtmlOptions.RestrictHttpLinksToBaseOrigin = true;
        o.HtmlOptions.RestrictHttpImagesToBaseOrigin = true;
        o.HtmlOptions.BlockExternalHttpImages = true;

        o.HtmlOptions.ImagesLoadingLazy = true;
        o.HtmlOptions.ImagesDecodingAsync = true;
        o.HtmlOptions.ImagesReferrerPolicy = "no-referrer";

        // Common chat UX helpers
        o.EnableCodeCopyButtons = true;
        o.EnableTableCopyButtons = true;

        // Guardrails to keep WebView hosts responsive under streaming/extreme outputs.
        o.MaxMarkdownChars = 500_000;
        o.MaxBodyHtmlBytes = 5_000_000;
        o.MarkdownOverflowHandling = OverflowHandling.Truncate;
        o.BodyHtmlOverflowHandling = OverflowHandling.RenderError;

        return o;
    }

    /// <summary>
    /// Strict preset for untrusted chat messages, with optional client-side renderers disabled.
    /// This disables Mermaid/Chart/Math/Prism and the copy-button UX helpers to minimize script usage in the shell.
    /// </summary>
    public static MarkdownRendererOptions CreateChatStrictMinimal(string? baseHref = null) {
        var o = CreateChatStrict(baseHref);
        o.EnableCodeCopyButtons = false;
        o.EnableTableCopyButtons = false;

        o.Mermaid.Enabled = false;
        o.Chart.Enabled = false;
        o.Math.Enabled = false;
        if (o.HtmlOptions.Prism != null) o.HtmlOptions.Prism.Enabled = false;

        return o;
    }

    /// <summary>
    /// Relaxed preset for trusted/controlled content rendered in a WebView.
    /// - Allows HTML parsing but sanitizes raw HTML blocks (very conservative allowlist)
    /// - Allows external HTTP(S) images (unless further restricted via host/origin allowlists)
    /// </summary>
    public static MarkdownRendererOptions CreateChatRelaxed(string? baseHref = null) {
        var o = CreateChatStrict(baseHref);

        o.ReaderOptions.HtmlBlocks = true;
        o.ReaderOptions.InlineHtml = true;

        o.HtmlOptions.RawHtmlHandling = RawHtmlHandling.Sanitize;
        o.HtmlOptions.BlockExternalHttpImages = false;

        // In relaxed mode, don't suppress cross-origin HTTP(S) by default; let the host opt-in.
        o.HtmlOptions.RestrictHttpLinksToBaseOrigin = false;
        o.HtmlOptions.RestrictHttpImagesToBaseOrigin = false;

        return o;
    }
}
