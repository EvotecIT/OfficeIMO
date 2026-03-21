namespace OfficeIMO.Markdown;

/// <summary>
/// Footnote definition block, e.g., [^id]: content.
/// </summary>
public sealed class FootnoteDefinitionBlock : MarkdownBlock, IMarkdownBlock, IChildMarkdownBlockContainer, ISyntaxChildrenMarkdownBlock, IOwnedSyntaxChildrenMarkdownBlock, ISyntaxMarkdownBlock, IFootnoteSectionMarkdownBlock {
    /// <summary>Footnote label (identifier without the leading ^).</summary>
    public string Label { get; }
    private readonly string _fallbackText;
    private readonly IReadOnlyList<IMarkdownBlock> _blocks;
    private readonly IReadOnlyList<InlineSequence> _paragraphs;
    private readonly IReadOnlyList<ParagraphBlock> _paragraphBlocks;

    private readonly struct FootnoteContentViews(
        IReadOnlyList<IMarkdownBlock> blocks,
        IReadOnlyList<ParagraphBlock> paragraphBlocks,
        IReadOnlyList<InlineSequence> paragraphs) {
        public IReadOnlyList<IMarkdownBlock> Blocks { get; } = blocks ?? Array.Empty<IMarkdownBlock>();
        public IReadOnlyList<ParagraphBlock> ParagraphBlocks { get; } = paragraphBlocks ?? Array.Empty<ParagraphBlock>();
        public IReadOnlyList<InlineSequence> Paragraphs { get; } = paragraphs ?? Array.Empty<InlineSequence>();
    }

    /// <summary>Footnote text content. When parsed child blocks are available, this is derived from them.</summary>
    public string Text => _blocks.Count > 0 ? RenderBlocksAsText(_blocks) : _fallbackText;
    /// <summary>
     /// Parsed child blocks of the footnote definition.
     /// Paragraph-only footnotes keep this aligned with <see cref="ParagraphBlocks"/>, while richer
     /// producers can preserve headings, code blocks, and other block types here.
     /// </summary>
    public IReadOnlyList<IMarkdownBlock> Blocks => _blocks;
    /// <summary>
     /// Parsed paragraphs of the footnote definition (when created by the reader).
     /// When empty, renderers may fall back to parsing <see cref="Text"/> as a single inline sequence.
     /// </summary>
    public IReadOnlyList<InlineSequence> Paragraphs => _paragraphs;
    /// <summary>
     /// Parsed paragraph blocks of the footnote definition (when created by the reader).
     /// This exposes footnote content as owned block children for AST-style consumers.
     /// </summary>
    public IReadOnlyList<ParagraphBlock> ParagraphBlocks => _paragraphBlocks;
    internal IReadOnlyList<MarkdownSyntaxNode>? SyntaxChildren { get; }
    string IFootnoteSectionMarkdownBlock.FootnoteLabel => Label;

    private FootnoteDefinitionBlock(
        string label,
        string fallbackText,
        FootnoteContentViews content,
        IReadOnlyList<MarkdownSyntaxNode>? syntaxChildren) {
        Label = label ?? string.Empty;
        _fallbackText = fallbackText ?? string.Empty;
        _blocks = content.Blocks;
        _paragraphBlocks = content.ParagraphBlocks;
        _paragraphs = content.Paragraphs;
        SyntaxChildren = syntaxChildren;
    }

    /// <summary>Create a new footnote definition.</summary>
    /// <param name="label">Identifier used by references.</param>
    /// <param name="text">Definition text.</param>
    public FootnoteDefinitionBlock(string label, string text)
        : this(
            label,
            text,
            new FootnoteContentViews(
                Array.Empty<IMarkdownBlock>(),
                Array.Empty<ParagraphBlock>(),
                Array.Empty<InlineSequence>()),
            syntaxChildren: null) {
    }

    internal FootnoteDefinitionBlock(string label, string text, IReadOnlyList<InlineSequence> paragraphs)
        : this(label, text, paragraphs, syntaxChildren: null) {
    }

    internal FootnoteDefinitionBlock(string label, string text, IReadOnlyList<ParagraphBlock> paragraphBlocks)
        : this(label, text, paragraphBlocks, syntaxChildren: null) {
    }

    internal FootnoteDefinitionBlock(string label, string text, IReadOnlyList<InlineSequence> paragraphs, IReadOnlyList<MarkdownSyntaxNode>? syntaxChildren)
        : this(
            label,
            text,
            CreateContentViewsFromParagraphs(paragraphs),
            syntaxChildren) {
    }

    internal FootnoteDefinitionBlock(string label, string text, IReadOnlyList<ParagraphBlock> paragraphBlocks, IReadOnlyList<MarkdownSyntaxNode>? syntaxChildren)
        : this(
            label,
            text,
            CreateContentViewsFromParagraphBlocks(paragraphBlocks),
            syntaxChildren) {
    }

    internal FootnoteDefinitionBlock(string label, string text, IReadOnlyList<IMarkdownBlock> blocks, IReadOnlyList<MarkdownSyntaxNode>? syntaxChildren)
        : this(
            label,
            text,
            CreateContentViewsFromBlocks(blocks),
            syntaxChildren) {
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
            ((IOwnedSyntaxChildrenMarkdownBlock)this).BuildOwnedSyntaxChildren(),
            this);

    private string RenderMarkdown() {
        var blocks = GetBlocksForRender();
        if (blocks.Count == 0) {
            return $"[^{Label}]: {Text}";
        }

        var sb = new System.Text.StringBuilder();
        sb.Append("[^").Append(Label).Append("]:");
        if (blocks[0] is ParagraphBlock) {
            sb.Append(' ');
            AppendIndentedBlockMarkdown(sb, blocks[0], firstBlock: true);
        } else {
            sb.Append('\n');
            AppendIndentedBlockMarkdown(sb, blocks[0], firstBlock: false);
        }

        for (int i = 1; i < blocks.Count; i++) {
            sb.Append("\n\n");
            AppendIndentedBlockMarkdown(sb, blocks[i], firstBlock: false);
        }

        return sb.ToString();
    }

    private static IReadOnlyList<ParagraphBlock> CreateParagraphBlocks(IReadOnlyList<InlineSequence> paragraphs) {
        if (paragraphs == null || paragraphs.Count == 0) {
            return Array.Empty<ParagraphBlock>();
        }

        var blocks = new List<ParagraphBlock>(paragraphs.Count);
        for (int i = 0; i < paragraphs.Count; i++) {
            blocks.Add(new ParagraphBlock(paragraphs[i] ?? new InlineSequence()));
        }
        return blocks;
    }

    private static IReadOnlyList<InlineSequence> CreateParagraphInlines(IReadOnlyList<ParagraphBlock> paragraphBlocks) {
        if (paragraphBlocks == null || paragraphBlocks.Count == 0) {
            return Array.Empty<InlineSequence>();
        }

        var paragraphs = new List<InlineSequence>(paragraphBlocks.Count);
        for (int i = 0; i < paragraphBlocks.Count; i++) {
            paragraphs.Add(paragraphBlocks[i]?.Inlines ?? new InlineSequence());
        }
        return paragraphs;
    }

    private IReadOnlyList<IMarkdownBlock> GetBlocksForRender() {
        if (_blocks.Count > 0) {
            return _blocks;
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
            return Array.Empty<IMarkdownBlock>();
        }

        var blocks = new List<IMarkdownBlock>(paragraphBlocks.Count);
        for (int i = 0; i < paragraphBlocks.Count; i++) {
            blocks.Add(paragraphBlocks[i] ?? new ParagraphBlock(new InlineSequence()));
        }

        return blocks;
    }

    private static IReadOnlyList<ParagraphBlock> CreateParagraphBlocks(IReadOnlyList<IMarkdownBlock> blocks) {
        if (blocks == null || blocks.Count == 0) {
            return Array.Empty<ParagraphBlock>();
        }

        var paragraphs = new List<ParagraphBlock>();
        for (int i = 0; i < blocks.Count; i++) {
            if (blocks[i] is ParagraphBlock paragraph) {
                paragraphs.Add(paragraph);
            }
        }

        return paragraphs;
    }

    private static IReadOnlyList<IMarkdownBlock> CopyBlocks(IReadOnlyList<IMarkdownBlock>? blocks) {
        if (blocks == null || blocks.Count == 0) {
            return Array.Empty<IMarkdownBlock>();
        }

        var copy = new IMarkdownBlock[blocks.Count];
        for (int i = 0; i < blocks.Count; i++) {
            copy[i] = blocks[i];
        }

        return copy;
    }

    private static FootnoteContentViews CreateContentViewsFromParagraphs(IReadOnlyList<InlineSequence>? paragraphs) {
        var copiedParagraphs = CopyParagraphs(paragraphs);
        var paragraphBlocks = CreateParagraphBlocks(copiedParagraphs);
        return new FootnoteContentViews(
            CreateBlockList(paragraphBlocks),
            paragraphBlocks,
            copiedParagraphs);
    }

    private static FootnoteContentViews CreateContentViewsFromParagraphBlocks(IReadOnlyList<ParagraphBlock>? paragraphBlocks) {
        var copiedParagraphBlocks = CopyParagraphBlocks(paragraphBlocks);
        return new FootnoteContentViews(
            CreateBlockList(copiedParagraphBlocks),
            copiedParagraphBlocks,
            CreateParagraphInlines(copiedParagraphBlocks));
    }

    private static FootnoteContentViews CreateContentViewsFromBlocks(IReadOnlyList<IMarkdownBlock>? blocks) {
        var copiedBlocks = CopyBlocks(blocks);
        var paragraphBlocks = CreateParagraphBlocks(copiedBlocks);
        return new FootnoteContentViews(
            copiedBlocks,
            paragraphBlocks,
            CreateParagraphInlines(paragraphBlocks));
    }

    private static IReadOnlyList<ParagraphBlock> CopyParagraphBlocks(IReadOnlyList<ParagraphBlock>? paragraphBlocks) {
        if (paragraphBlocks == null || paragraphBlocks.Count == 0) {
            return Array.Empty<ParagraphBlock>();
        }

        var copy = new ParagraphBlock[paragraphBlocks.Count];
        for (int i = 0; i < paragraphBlocks.Count; i++) {
            copy[i] = paragraphBlocks[i];
        }

        return copy;
    }

    private static IReadOnlyList<InlineSequence> CopyParagraphs(IReadOnlyList<InlineSequence>? paragraphs) {
        if (paragraphs == null || paragraphs.Count == 0) {
            return Array.Empty<InlineSequence>();
        }

        var copy = new InlineSequence[paragraphs.Count];
        for (int i = 0; i < paragraphs.Count; i++) {
            copy[i] = paragraphs[i] ?? new InlineSequence();
        }

        return copy;
    }

    private static string RenderBlocksAsText(IReadOnlyList<IMarkdownBlock> blocks) {
        if (blocks == null || blocks.Count == 0) {
            return string.Empty;
        }

        var sb = new System.Text.StringBuilder();
        for (int i = 0; i < blocks.Count; i++) {
            if (blocks[i] == null) {
                continue;
            }

            if (sb.Length > 0) {
                sb.Append("\n\n");
            }

            sb.Append((blocks[i].RenderMarkdown() ?? string.Empty)
                .Replace("\r\n", "\n")
                .Replace('\r', '\n'));
        }

        return sb.ToString();
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
