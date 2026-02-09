namespace OfficeIMO.Markdown;

/// <summary>
/// Italic/emphasis that contains nested inline nodes.
/// Used by the reader so nested markup can be represented without changing the fluent builder API.
/// </summary>
public sealed class ItalicSequenceInline {
    /// <summary>Inline content.</summary>
    public InlineSequence Inlines { get; }

    /// <summary>Creates an italic inline with nested inline content.</summary>
    public ItalicSequenceInline(InlineSequence inlines) {
        Inlines = inlines ?? new InlineSequence();
    }

    internal string RenderMarkdown() => "*" + Inlines.RenderMarkdown() + "*";
    internal string RenderHtml() => "<em>" + Inlines.RenderHtml() + "</em>";
}

