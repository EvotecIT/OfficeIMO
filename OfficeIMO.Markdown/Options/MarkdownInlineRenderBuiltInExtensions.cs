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
    /// <summary>Stable registration name for the CommonMark inserted markdown fallback.</summary>
    public const string CommonMarkInsertedMarkdownName = "CommonMark.Inserted.Markdown";
    /// <summary>Stable registration name for the CommonMark superscript markdown fallback.</summary>
    public const string CommonMarkSuperscriptMarkdownName = "CommonMark.Superscript.Markdown";
    /// <summary>Stable registration name for the CommonMark subscript markdown fallback.</summary>
    public const string CommonMarkSubscriptMarkdownName = "CommonMark.Subscript.Markdown";

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
        AddIfMissing(options.InlineRenderExtensions, CommonMarkInsertedMarkdownName, typeof(InsertedInline), RenderHtmlFallback);
        AddIfMissing(options.InlineRenderExtensions, CommonMarkInsertedMarkdownName + ".Sequence", typeof(InsertedSequenceInline), RenderHtmlFallback);
        AddIfMissing(options.InlineRenderExtensions, CommonMarkSuperscriptMarkdownName, typeof(SuperscriptInline), RenderHtmlFallback);
        AddIfMissing(options.InlineRenderExtensions, CommonMarkSuperscriptMarkdownName + ".Sequence", typeof(SuperscriptSequenceInline), RenderHtmlFallback);
        AddIfMissing(options.InlineRenderExtensions, CommonMarkSubscriptMarkdownName, typeof(SubscriptInline), RenderHtmlFallback);
        AddIfMissing(options.InlineRenderExtensions, CommonMarkSubscriptMarkdownName + ".Sequence", typeof(SubscriptSequenceInline), RenderHtmlFallback);
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
            InsertedInline inserted => inserted.RenderHtml(),
            InsertedSequenceInline inserted => inserted.RenderHtml(),
            SuperscriptInline superscript => superscript.RenderHtml(),
            SuperscriptSequenceInline superscript => superscript.RenderHtml(),
            SubscriptInline subscript => subscript.RenderHtml(),
            SubscriptSequenceInline subscript => subscript.RenderHtml(),
            _ => null
        };

        return html == null
            ? null
            : MarkdownInlineAttributeRenderer.RenderHtml(inline, html, null);
    }
}
