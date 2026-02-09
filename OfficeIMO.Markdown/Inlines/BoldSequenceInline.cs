namespace OfficeIMO.Markdown;

/// <summary>
/// Bold/strong emphasis that contains nested inline nodes.
/// Used by the reader so nested markup can be represented without changing the fluent builder API.
/// </summary>
public sealed class BoldSequenceInline {
    /// <summary>Inline content.</summary>
    public InlineSequence Inlines { get; }

    /// <summary>Creates a bold inline with nested inline content.</summary>
    public BoldSequenceInline(InlineSequence inlines) {
        Inlines = inlines ?? new InlineSequence();
    }

    internal string RenderMarkdown() => "**" + Inlines.RenderMarkdown() + "**";
    internal string RenderHtml() => "<strong>" + Inlines.RenderHtml() + "</strong>";
}

