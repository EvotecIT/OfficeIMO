namespace OfficeIMO.Markdown;

/// <summary>
/// Footnote definition block, e.g., [^id]: content.
/// </summary>
public sealed class FootnoteDefinitionBlock : IMarkdownBlock, ISyntaxChildrenMarkdownBlock, ISyntaxMarkdownBlock {
    /// <summary>Footnote label (identifier without the leading ^).</summary>
    public string Label { get; }
    /// <summary>Footnote text content.</summary>
    public string Text { get; }
    /// <summary>
    /// Parsed paragraphs of the footnote definition (when created by the reader).
    /// When empty, renderers may fall back to parsing <see cref="Text"/> as a single inline sequence.
    /// </summary>
    public IReadOnlyList<InlineSequence> Paragraphs { get; }
    internal IReadOnlyList<MarkdownSyntaxNode>? SyntaxChildren { get; }
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
        SyntaxChildren = null;
    }

    internal FootnoteDefinitionBlock(string label, string text, IReadOnlyList<InlineSequence> paragraphs, IReadOnlyList<MarkdownSyntaxNode>? syntaxChildren) {
        Label = label ?? string.Empty;
        Text = text ?? string.Empty;
        Paragraphs = paragraphs ?? new List<InlineSequence>();
        SyntaxChildren = syntaxChildren;
    }
    string IMarkdownBlock.RenderMarkdown() => $"[^{Label}]: {Text}";
    string IMarkdownBlock.RenderHtml() {
        // Standalone rendering; HtmlRenderer aggregates footnotes into a dedicated section.
        var encLabel = System.Net.WebUtility.HtmlEncode(Label);
        var paragraphs = Paragraphs;
        if (paragraphs != null && paragraphs.Count > 0) {
            var sb = new System.Text.StringBuilder();
            for (int i = 0; i < paragraphs.Count; i++) {
                var paragraph = paragraphs[i] ?? new InlineSequence();
                sb.Append("<p id=\"fn:").Append(encLabel).Append("\"><sup>").Append(encLabel).Append("</sup> ")
                    .Append(paragraph.RenderHtml());
                if (i == paragraphs.Count - 1) {
                    sb.Append(" <a class=\"footnote-backref\" href=\"#fnref:").Append(encLabel).Append("\" aria-label=\"Back to reference\">&#8617;</a>");
                }
                sb.Append("</p>");
            }

            return sb.ToString();
        }

        var inlines = MarkdownReader.ParseInlineText(Text);
        return $"<p id=\"fn:{encLabel}\"><sup>{encLabel}</sup> {inlines.RenderHtml()} <a class=\"footnote-backref\" href=\"#fnref:{encLabel}\" aria-label=\"Back to reference\">&#8617;</a></p>";
    }

    IReadOnlyList<MarkdownSyntaxNode>? ISyntaxChildrenMarkdownBlock.ProvidedSyntaxChildren => SyntaxChildren;

    internal IReadOnlyList<MarkdownSyntaxNode> BuildSyntaxChildren() {
        if (SyntaxChildren != null && SyntaxChildren.Count > 0) {
            return SyntaxChildren;
        }

        if (Paragraphs.Count == 0) {
            return Array.Empty<MarkdownSyntaxNode>();
        }

        var nodes = new List<MarkdownSyntaxNode>(Paragraphs.Count);
        for (int i = 0; i < Paragraphs.Count; i++) {
            nodes.Add(new MarkdownSyntaxNode(MarkdownSyntaxKind.Paragraph, literal: Paragraphs[i].RenderMarkdown()));
        }
        return nodes;
    }

    MarkdownSyntaxNode ISyntaxMarkdownBlock.BuildSyntaxNode(MarkdownSourceSpan? span) =>
        MarkdownBlockSyntaxBuilder.BuildFootnoteBlock(this, span);
}
