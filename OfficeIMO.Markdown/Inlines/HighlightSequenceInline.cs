namespace OfficeIMO.Markdown;

/// <summary>
/// Highlighted inline content that can contain nested inline nodes.
/// Used by the reader so nested markup can be represented without flattening formatting.
/// </summary>
public sealed class HighlightSequenceInline {
    /// <summary>Inline content.</summary>
    public InlineSequence Inlines { get; }

    /// <summary>Creates a highlighted inline with nested inline content.</summary>
    public HighlightSequenceInline(InlineSequence inlines) {
        Inlines = inlines ?? new InlineSequence();
    }

    internal string RenderMarkdown() => "==" + Inlines.RenderMarkdown() + "==";
    internal string RenderHtml() => "<mark>" + Inlines.RenderHtml() + "</mark>";
}
