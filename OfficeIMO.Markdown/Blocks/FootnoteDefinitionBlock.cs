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
    /// Parsed child blocks of the footnote definition.
    /// Paragraph-only footnotes keep this aligned with <see cref="ParagraphBlocks"/>, while richer
    /// producers can preserve headings, code blocks, and other block types here.
    /// </summary>
    public IReadOnlyList<IMarkdownBlock> Blocks { get; }
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
        Blocks = new List<IMarkdownBlock>();
        Paragraphs = new List<InlineSequence>();
        ParagraphBlocks = new List<ParagraphBlock>();
    }

    internal FootnoteDefinitionBlock(string label, string text, IReadOnlyList<InlineSequence> paragraphs) {
        Label = label ?? string.Empty;
        Text = text ?? string.Empty;
        Paragraphs = paragraphs ?? new List<InlineSequence>();
        ParagraphBlocks = CreateParagraphBlocks(Paragraphs);
        Blocks = CreateBlockList(ParagraphBlocks);
        SyntaxChildren = null;
    }

    internal FootnoteDefinitionBlock(string label, string text, IReadOnlyList<ParagraphBlock> paragraphBlocks) {
        Label = label ?? string.Empty;
        Text = text ?? string.Empty;
        ParagraphBlocks = paragraphBlocks ?? new List<ParagraphBlock>();
        Paragraphs = CreateParagraphInlines(ParagraphBlocks);
        Blocks = CreateBlockList(ParagraphBlocks);
        SyntaxChildren = null;
    }

    internal FootnoteDefinitionBlock(string label, string text, IReadOnlyList<InlineSequence> paragraphs, IReadOnlyList<MarkdownSyntaxNode>? syntaxChildren) {
        Label = label ?? string.Empty;
        Text = text ?? string.Empty;
        Paragraphs = paragraphs ?? new List<InlineSequence>();
        ParagraphBlocks = CreateParagraphBlocks(Paragraphs);
        Blocks = CreateBlockList(ParagraphBlocks);
        SyntaxChildren = syntaxChildren;
    }

    internal FootnoteDefinitionBlock(string label, string text, IReadOnlyList<ParagraphBlock> paragraphBlocks, IReadOnlyList<MarkdownSyntaxNode>? syntaxChildren) {
        Label = label ?? string.Empty;
        Text = text ?? string.Empty;
        ParagraphBlocks = paragraphBlocks ?? new List<ParagraphBlock>();
        Paragraphs = CreateParagraphInlines(ParagraphBlocks);
        Blocks = CreateBlockList(ParagraphBlocks);
        SyntaxChildren = syntaxChildren;
    }

    internal FootnoteDefinitionBlock(string label, string text, IReadOnlyList<IMarkdownBlock> blocks, IReadOnlyList<MarkdownSyntaxNode>? syntaxChildren) {
        Label = label ?? string.Empty;
        Text = text ?? string.Empty;
        Blocks = blocks ?? new List<IMarkdownBlock>();
        ParagraphBlocks = CreateParagraphBlocks(Blocks);
        Paragraphs = CreateParagraphInlines(ParagraphBlocks);
        SyntaxChildren = syntaxChildren;
    }
    string IMarkdownBlock.RenderMarkdown() => RenderMarkdown();
    string IMarkdownBlock.RenderHtml() {
        // Standalone rendering; HtmlRenderer aggregates footnotes into a dedicated section.
        var encLabel = System.Net.WebUtility.HtmlEncode(Label);
        var blocks = GetBlocksForRender();
        if (AreAllParagraphBlocks(blocks)) {
            var paragraphs = CreateParagraphBlocks(blocks);
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

        var mixed = new System.Text.StringBuilder();
        mixed.Append("<div id=\"fn:").Append(encLabel).Append("\"><p><sup>").Append(encLabel).Append("</sup></p>");
        for (int i = 0; i < blocks.Count; i++) {
            var block = blocks[i];
            if (block == null) {
                continue;
            }

            mixed.Append(block.RenderHtml());
        }

        mixed.Append("<p><a class=\"footnote-backref\" href=\"#fnref:").Append(encLabel).Append("\" aria-label=\"Back to reference\">&#8617;</a></p></div>");
        return mixed.ToString();
    }

    string IFootnoteSectionMarkdownBlock.RenderFootnoteSectionItemHtml() {
        var label = Label ?? string.Empty;
        if (label.Length == 0) {
            return string.Empty;
        }

        var encLabel = System.Net.WebUtility.HtmlEncode(label);
        var blocks = GetBlocksForRender();

        var sb = new System.Text.StringBuilder();
        sb.Append("<li id=\"fn:").Append(encLabel).Append("\">");

        if (AreAllParagraphBlocks(blocks)) {
            var paragraphs = CreateParagraphBlocks(blocks);
            for (int i = 0; i < paragraphs.Count; i++) {
                var paragraph = paragraphs[i] ?? new ParagraphBlock(new InlineSequence());
                sb.Append("<p>").Append(paragraph.Inlines.RenderHtml());
                if (i == paragraphs.Count - 1) {
                    sb.Append(" <a class=\"footnote-backref\" href=\"#fnref:").Append(encLabel).Append("\" aria-label=\"Back to reference\">&#8617;</a>");
                }
                sb.Append("</p>");
            }
        } else {
            for (int i = 0; i < blocks.Count; i++) {
                var block = blocks[i];
                if (block == null) {
                    continue;
                }

                sb.Append(block.RenderHtml());
            }

            sb.Append("<p><a class=\"footnote-backref\" href=\"#fnref:").Append(encLabel).Append("\" aria-label=\"Back to reference\">&#8617;</a></p>");
        }

        sb.Append("</li>");
        return sb.ToString();
    }

    IReadOnlyList<IMarkdownBlock> IChildMarkdownBlockContainer.ChildBlocks => Blocks;
    IReadOnlyList<MarkdownSyntaxNode>? ISyntaxChildrenMarkdownBlock.ProvidedSyntaxChildren => SyntaxChildren;

    IReadOnlyList<MarkdownSyntaxNode> IOwnedSyntaxChildrenMarkdownBlock.BuildOwnedSyntaxChildren() {
        if (SyntaxChildren != null && SyntaxChildren.Count > 0) {
            return SyntaxChildren;
        }

        if (Blocks.Count == 0) {
            return Array.Empty<MarkdownSyntaxNode>();
        }

        return MarkdownBlockSyntaxBuilder.BuildChildSyntaxNodes(Blocks);
    }

    MarkdownSyntaxNode ISyntaxMarkdownBlock.BuildSyntaxNode(MarkdownSourceSpan? span) =>
        new MarkdownSyntaxNode(
            MarkdownSyntaxKind.FootnoteDefinition,
            span,
            Label,
            ((IOwnedSyntaxChildrenMarkdownBlock)this).BuildOwnedSyntaxChildren());

    private string RenderMarkdown() {
        var blocks = GetBlocksForRender();
        if (blocks.Count == 0) {
            return $"[^{Label}]: {Text}";
        }

        var sb = new System.Text.StringBuilder();
        sb.Append("[^").Append(Label).Append("]: ");
        AppendIndentedBlockMarkdown(sb, blocks[0], firstBlock: true);

        for (int i = 1; i < blocks.Count; i++) {
            sb.Append("\n\n");
            AppendIndentedBlockMarkdown(sb, blocks[i], firstBlock: false);
        }

        return sb.ToString();
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

    private IReadOnlyList<IMarkdownBlock> GetBlocksForRender() {
        if (Blocks.Count > 0) {
            return Blocks;
        }

        return new List<IMarkdownBlock> { new ParagraphBlock(MarkdownReader.ParseInlineText(Text)) };
    }

    private static bool AreAllParagraphBlocks(IReadOnlyList<IMarkdownBlock> blocks) {
        if (blocks == null || blocks.Count == 0) {
            return false;
        }

        for (int i = 0; i < blocks.Count; i++) {
            if (blocks[i] is not ParagraphBlock) {
                return false;
            }
        }

        return true;
    }

    private static IReadOnlyList<IMarkdownBlock> CreateBlockList(IReadOnlyList<ParagraphBlock> paragraphBlocks) {
        if (paragraphBlocks == null || paragraphBlocks.Count == 0) {
            return new List<IMarkdownBlock>();
        }

        var blocks = new List<IMarkdownBlock>(paragraphBlocks.Count);
        for (int i = 0; i < paragraphBlocks.Count; i++) {
            blocks.Add(paragraphBlocks[i] ?? new ParagraphBlock(new InlineSequence()));
        }

        return blocks;
    }

    private static IReadOnlyList<ParagraphBlock> CreateParagraphBlocks(IReadOnlyList<IMarkdownBlock> blocks) {
        if (blocks == null || blocks.Count == 0) {
            return new List<ParagraphBlock>();
        }

        var paragraphs = new List<ParagraphBlock>();
        for (int i = 0; i < blocks.Count; i++) {
            if (blocks[i] is ParagraphBlock paragraph) {
                paragraphs.Add(paragraph);
            }
        }

        return paragraphs;
    }

    private static void AppendIndentedBlockMarkdown(System.Text.StringBuilder sb, IMarkdownBlock block, bool firstBlock) {
        string rendered = (block?.RenderMarkdown() ?? string.Empty)
            .Replace("\r\n", "\n")
            .Replace('\r', '\n');
        string[] lines = rendered.Split('\n');

        for (int i = 0; i < lines.Length; i++) {
            if (i > 0) {
                sb.Append('\n');
                sb.Append("  ");
            } else if (!firstBlock) {
                sb.Append("  ");
            }

            sb.Append(lines[i]);
        }
    }
}
