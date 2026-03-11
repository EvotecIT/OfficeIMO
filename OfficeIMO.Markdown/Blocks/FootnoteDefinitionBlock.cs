namespace OfficeIMO.Markdown;

/// <summary>
/// Footnote definition block, e.g., [^id]: content.
/// </summary>
public sealed class FootnoteDefinitionBlock : IMarkdownBlock, IChildMarkdownBlockContainer, ISyntaxChildrenMarkdownBlock, IOwnedSyntaxChildrenMarkdownBlock, ISyntaxMarkdownBlock, IFootnoteSectionMarkdownBlock {
    /// <summary>Footnote label (identifier without the leading ^).</summary>
    public string Label { get; }
    /// <summary>Footnote text content.</summary>
    public string Text { get; }
    /// <summary>
    /// Parsed paragraphs of the footnote definition (when created by the reader).
    /// When empty, renderers may fall back to parsing <see cref="Text"/> as a single inline sequence.
    /// </summary>
    public IReadOnlyList<InlineSequence> Paragraphs { get; }
    /// <summary>
    /// Parsed paragraph blocks of the footnote definition (when created by the reader).
    /// This exposes footnote content as owned block children for AST-style consumers.
    /// </summary>
    public IReadOnlyList<ParagraphBlock> ParagraphBlocks { get; }
    internal IReadOnlyList<MarkdownSyntaxNode>? SyntaxChildren { get; }
    string IFootnoteSectionMarkdownBlock.FootnoteLabel => Label;
    /// <summary>Create a new footnote definition.</summary>
    /// <param name="label">Identifier used by references.</param>
    /// <param name="text">Definition text.</param>
    public FootnoteDefinitionBlock(string label, string text) {
        Label = label ?? string.Empty;
        Text = text ?? string.Empty;
        Paragraphs = new List<InlineSequence>();
        ParagraphBlocks = new List<ParagraphBlock>();
    }

    internal FootnoteDefinitionBlock(string label, string text, IReadOnlyList<InlineSequence> paragraphs) {
        Label = label ?? string.Empty;
        Text = text ?? string.Empty;
        Paragraphs = paragraphs ?? new List<InlineSequence>();
        ParagraphBlocks = CreateParagraphBlocks(Paragraphs);
        SyntaxChildren = null;
    }

    internal FootnoteDefinitionBlock(string label, string text, IReadOnlyList<ParagraphBlock> paragraphBlocks) {
        Label = label ?? string.Empty;
        Text = text ?? string.Empty;
        ParagraphBlocks = paragraphBlocks ?? new List<ParagraphBlock>();
        Paragraphs = CreateParagraphInlines(ParagraphBlocks);
        SyntaxChildren = null;
    }

    internal FootnoteDefinitionBlock(string label, string text, IReadOnlyList<InlineSequence> paragraphs, IReadOnlyList<MarkdownSyntaxNode>? syntaxChildren) {
        Label = label ?? string.Empty;
        Text = text ?? string.Empty;
        Paragraphs = paragraphs ?? new List<InlineSequence>();
        ParagraphBlocks = CreateParagraphBlocks(Paragraphs);
        SyntaxChildren = syntaxChildren;
    }

    internal FootnoteDefinitionBlock(string label, string text, IReadOnlyList<ParagraphBlock> paragraphBlocks, IReadOnlyList<MarkdownSyntaxNode>? syntaxChildren) {
        Label = label ?? string.Empty;
        Text = text ?? string.Empty;
        ParagraphBlocks = paragraphBlocks ?? new List<ParagraphBlock>();
        Paragraphs = CreateParagraphInlines(ParagraphBlocks);
        SyntaxChildren = syntaxChildren;
    }
    string IMarkdownBlock.RenderMarkdown() => $"[^{Label}]: {Text}";
    string IMarkdownBlock.RenderHtml() {
        // Standalone rendering; HtmlRenderer aggregates footnotes into a dedicated section.
        var encLabel = System.Net.WebUtility.HtmlEncode(Label);
        var paragraphs = GetParagraphBlocksForRender();
        if (paragraphs.Count > 0) {
            var sb = new System.Text.StringBuilder();
            for (int i = 0; i < paragraphs.Count; i++) {
                var paragraph = paragraphs[i] ?? new ParagraphBlock(new InlineSequence());
                sb.Append("<p id=\"fn:").Append(encLabel).Append("\"><sup>").Append(encLabel).Append("</sup> ")
                    .Append(paragraph.Inlines.RenderHtml());
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

    string IFootnoteSectionMarkdownBlock.RenderFootnoteSectionItemHtml() {
        var label = Label ?? string.Empty;
        if (label.Length == 0) {
            return string.Empty;
        }

        var encLabel = System.Net.WebUtility.HtmlEncode(label);
        var paragraphs = GetParagraphBlocksForRender();

        var sb = new System.Text.StringBuilder();
        sb.Append("<li id=\"fn:").Append(encLabel).Append("\">");

        for (int i = 0; i < paragraphs.Count; i++) {
            var paragraph = paragraphs[i] ?? new ParagraphBlock(new InlineSequence());
            sb.Append("<p>").Append(paragraph.Inlines.RenderHtml());
            if (i == paragraphs.Count - 1) {
                sb.Append(" <a class=\"footnote-backref\" href=\"#fnref:").Append(encLabel).Append("\" aria-label=\"Back to reference\">&#8617;</a>");
            }
            sb.Append("</p>");
        }

        sb.Append("</li>");
        return sb.ToString();
    }

    IReadOnlyList<IMarkdownBlock> IChildMarkdownBlockContainer.ChildBlocks => ParagraphBlocks;
    IReadOnlyList<MarkdownSyntaxNode>? ISyntaxChildrenMarkdownBlock.ProvidedSyntaxChildren => SyntaxChildren;

    IReadOnlyList<MarkdownSyntaxNode> IOwnedSyntaxChildrenMarkdownBlock.BuildOwnedSyntaxChildren() {
        if (SyntaxChildren != null && SyntaxChildren.Count > 0) {
            return SyntaxChildren;
        }

        if (ParagraphBlocks.Count == 0) {
            return Array.Empty<MarkdownSyntaxNode>();
        }

        return MarkdownBlockSyntaxBuilder.BuildChildSyntaxNodes(ParagraphBlocks);
    }

    MarkdownSyntaxNode ISyntaxMarkdownBlock.BuildSyntaxNode(MarkdownSourceSpan? span) =>
        new MarkdownSyntaxNode(
            MarkdownSyntaxKind.FootnoteDefinition,
            span,
            Label,
            ((IOwnedSyntaxChildrenMarkdownBlock)this).BuildOwnedSyntaxChildren());

    private IReadOnlyList<ParagraphBlock> GetParagraphBlocksForRender() {
        if (ParagraphBlocks.Count > 0) {
            return ParagraphBlocks;
        }

        return new List<ParagraphBlock> { new ParagraphBlock(MarkdownReader.ParseInlineText(Text)) };
    }

    private static IReadOnlyList<ParagraphBlock> CreateParagraphBlocks(IReadOnlyList<InlineSequence> paragraphs) {
        if (paragraphs == null || paragraphs.Count == 0) {
            return new List<ParagraphBlock>();
        }

        var blocks = new List<ParagraphBlock>(paragraphs.Count);
        for (int i = 0; i < paragraphs.Count; i++) {
            blocks.Add(new ParagraphBlock(paragraphs[i] ?? new InlineSequence()));
        }
        return blocks;
    }

    private static IReadOnlyList<InlineSequence> CreateParagraphInlines(IReadOnlyList<ParagraphBlock> paragraphBlocks) {
        if (paragraphBlocks == null || paragraphBlocks.Count == 0) {
            return new List<InlineSequence>();
        }

        var paragraphs = new List<InlineSequence>(paragraphBlocks.Count);
        for (int i = 0; i < paragraphBlocks.Count; i++) {
            paragraphs.Add(paragraphBlocks[i]?.Inlines ?? new InlineSequence());
        }
        return paragraphs;
    }
}
