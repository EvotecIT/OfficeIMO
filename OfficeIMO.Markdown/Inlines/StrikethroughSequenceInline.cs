namespace OfficeIMO.Markdown;

/// <summary>
/// Strikethrough that contains nested inline nodes.
/// Used by the reader so nested markup can be represented without changing the fluent builder API.
/// </summary>
public sealed class StrikethroughSequenceInline {
    /// <summary>Inline content.</summary>
    public InlineSequence Inlines { get; }

    /// <summary>Creates a strikethrough inline with nested inline content.</summary>
    public StrikethroughSequenceInline(InlineSequence inlines) {
        Inlines = inlines ?? new InlineSequence();
    }

    internal string RenderMarkdown() => "~~" + Inlines.RenderMarkdown() + "~~";
    internal string RenderHtml() => "<del>" + Inlines.RenderHtml() + "</del>";
}

