namespace OfficeIMO.Markdown;

/// <summary>
/// Footnote definition block, e.g., [^id]: content.
/// </summary>
public sealed class FootnoteDefinitionBlock : IMarkdownBlock {
    /// <summary>Footnote label (identifier without the leading ^).</summary>
    public string Label { get; }
    /// <summary>Footnote text content.</summary>
    public string Text { get; }
    /// <summary>
    /// Parsed paragraphs of the footnote definition (when created by the reader).
    /// When empty, renderers may fall back to parsing <see cref="Text"/> as a single inline sequence.
    /// </summary>
    public IReadOnlyList<InlineSequence> Paragraphs { get; }
    /// <summary>Create a new footnote definition.</summary>
    /// <param name="label">Identifier used by references.</param>
    /// <param name="text">Definition text.</param>
    public FootnoteDefinitionBlock(string label, string text) {
        Label = label ?? string.Empty;
        Text = text ?? string.Empty;
        Paragraphs = new List<InlineSequence>();
    }

    internal FootnoteDefinitionBlock(string label, string text, IReadOnlyList<InlineSequence> paragraphs) {
        Label = label ?? string.Empty;
        Text = text ?? string.Empty;
        Paragraphs = paragraphs ?? new List<InlineSequence>();
    }
    string IMarkdownBlock.RenderMarkdown() => $"[^{Label}]: {Text}";
    string IMarkdownBlock.RenderHtml() {
        // Standalone rendering; HtmlRenderer aggregates footnotes into a dedicated section.
        var encLabel = System.Net.WebUtility.HtmlEncode(Label);
        var inlines = MarkdownReader.ParseInlineText(Text);
        return $"<p id=\"fn:{encLabel}\"><sup>{encLabel}</sup> {inlines.RenderHtml()} <a class=\"footnote-backref\" href=\"#fnref:{encLabel}\" aria-label=\"Back to reference\">&#8617;</a></p>";
    }
}
