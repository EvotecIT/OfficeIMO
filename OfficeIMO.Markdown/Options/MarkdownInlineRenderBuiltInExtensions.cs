namespace OfficeIMO.Markdown;

/// <summary>
/// Built-in inline render extension registrations for portable or host-specific output shaping.
/// </summary>
public static class MarkdownInlineRenderBuiltInExtensions {
    /// <summary>Stable registration name for the CommonMark strikethrough markdown fallback.</summary>
    public const string CommonMarkStrikethroughMarkdownName = "CommonMark.Strikethrough.Markdown";

    /// <summary>Stable registration name for the CommonMark highlight markdown fallback.</summary>
    public const string CommonMarkHighlightMarkdownName = "CommonMark.Highlight.Markdown";
    /// <summary>Stable registration name for the CommonMark footnote-reference markdown fallback.</summary>
    public const string CommonMarkFootnoteReferenceMarkdownName = "CommonMark.FootnoteReference.Markdown";

    /// <summary>Adds CommonMark-compatible markdown fallbacks for GFM-only inline constructs.</summary>
    public static void AddCommonMarkGfmInlineMarkdownFallbacks(MarkdownWriteOptions options) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        AddIfMissing(options.InlineRenderExtensions, CommonMarkStrikethroughMarkdownName, typeof(StrikethroughInline), RenderHtmlFallback);
        AddIfMissing(options.InlineRenderExtensions, CommonMarkStrikethroughMarkdownName + ".Sequence", typeof(StrikethroughSequenceInline), RenderHtmlFallback);
        AddIfMissing(options.InlineRenderExtensions, CommonMarkHighlightMarkdownName, typeof(HighlightInline), RenderHtmlFallback);
        AddIfMissing(options.InlineRenderExtensions, CommonMarkHighlightMarkdownName + ".Sequence", typeof(HighlightSequenceInline), RenderHtmlFallback);
        AddIfMissing(options.InlineRenderExtensions, CommonMarkFootnoteReferenceMarkdownName, typeof(FootnoteRefInline), RenderHtmlFallback);
    }

    private static void AddIfMissing(
        List<MarkdownInlineMarkdownRenderExtension> extensions,
        string name,
        Type inlineType,
        MarkdownInlineMarkdownRenderer renderer) {
        if (extensions == null) {
            throw new ArgumentNullException(nameof(extensions));
        }

        if (extensions.Any(extension => string.Equals(extension.Name, name, StringComparison.OrdinalIgnoreCase))) {
            return;
        }

        extensions.Add(new MarkdownInlineMarkdownRenderExtension(name, inlineType, renderer));
    }

    private static string? RenderHtmlFallback(IMarkdownInline inline, MarkdownWriteOptions _) {
        string? html = inline switch {
            StrikethroughInline strikethrough => strikethrough.RenderHtml(),
            StrikethroughSequenceInline strikethrough => strikethrough.RenderHtml(),
            HighlightInline highlight => highlight.RenderHtml(),
            HighlightSequenceInline highlight => highlight.RenderHtml(),
            FootnoteRefInline footnote => footnote.RenderHtml(),
            _ => null
        };

        return html == null
            ? null
            : MarkdownInlineAttributeRenderer.RenderMarkdown(inline, html);
    }
}
