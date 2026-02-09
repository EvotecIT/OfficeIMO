namespace OfficeIMO.Markdown;

/// <summary>
/// Bold+italic emphasis that contains nested inline nodes.
/// Used by the reader so nested markup can be represented without changing the fluent builder API.
/// </summary>
public sealed class BoldItalicSequenceInline {
    /// <summary>Inline content.</summary>
    public InlineSequence Inlines { get; }

    /// <summary>Creates a bold+italic inline with nested inline content.</summary>
    public BoldItalicSequenceInline(InlineSequence inlines) {
        Inlines = inlines ?? new InlineSequence();
    }

    internal string RenderMarkdown() => "***" + Inlines.RenderMarkdown() + "***";
    internal string RenderHtml() => "<strong><em>" + Inlines.RenderHtml() + "</em></strong>";
}

