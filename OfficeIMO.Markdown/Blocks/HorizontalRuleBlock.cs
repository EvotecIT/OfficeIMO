namespace OfficeIMO.Markdown;

/// <summary>
/// Horizontal rule (thematic break). Rendered as --- in Markdown and <hr /> in HTML.
/// </summary>
public sealed class HorizontalRuleBlock : MarkdownBlock, IMarkdownBlock, ISyntaxMarkdownBlock {
    /// <summary>Source span of the thematic-break marker token when parsed from markdown.</summary>
    public MarkdownSourceSpan? MarkerSourceSpan { get; internal set; }
    /// <summary>Exact thematic-break marker text when parsed from markdown.</summary>
    public string? MarkerText { get; internal set; }

    string IMarkdownBlock.RenderMarkdown() => "---";
    string IMarkdownBlock.RenderHtml() => "<hr />";
    MarkdownSyntaxNode ISyntaxMarkdownBlock.BuildSyntaxNode(MarkdownSourceSpan? span) {
        var markerSpan = MarkerSourceSpan ?? span;
        var markerText = MarkerText ?? "---";
        var children = markerSpan.HasValue
            ? new[] { new MarkdownSyntaxNode(MarkdownSyntaxKind.ThematicBreakMarker, markerSpan.Value, markerText) }
            : Array.Empty<MarkdownSyntaxNode>();

        return new MarkdownSyntaxNode(MarkdownSyntaxKind.HorizontalRule, span ?? markerSpan, "---", children, this);
    }
}
