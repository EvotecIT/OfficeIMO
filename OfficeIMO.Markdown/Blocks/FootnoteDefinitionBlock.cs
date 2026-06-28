namespace OfficeIMO.Markdown;

/// <summary>
/// Footnote definition block, e.g., [^id]: content.
/// </summary>
public sealed class FootnoteDefinitionBlock : MarkdownBlock, IMarkdownBlock, IChildMarkdownBlockContainer, ISyntaxChildrenMarkdownBlock, IOwnedSyntaxChildrenMarkdownBlock, ISyntaxMarkdownBlock, IFootnoteSectionMarkdownBlock {
    /// <summary>Footnote label (identifier without the leading ^).</summary>
    public string Label { get; }
    /// <summary>Source span for the opening <c>[^</c> marker when parsed from markdown.</summary>
    public MarkdownSourceSpan? OpeningMarkerSourceSpan { get; internal set; }
    /// <summary>Source span for the footnote label token when parsed from markdown.</summary>
    public MarkdownSourceSpan? LabelSourceSpan { get; internal set; }
    /// <summary>Source span for the <c>]:</c> separator marker when parsed from markdown.</summary>
    public MarkdownSourceSpan? SeparatorMarkerSourceSpan { get; internal set; }
    private readonly string _fallbackText;
    private readonly string? _fallbackTextProjection;
    private readonly IReadOnlyList<IMarkdownBlock> _blocks;

    /// <summary>Footnote text content. When parsed child blocks are available, this is derived from them.</summary>
    public string Text {
        get {
            if (_blocks.Count == 0) {
                return _fallbackText;
            }

            if (string.IsNullOrEmpty(_fallbackText)) {
                return RenderBlocksAsText(_blocks);
            }

            var projectedText = RenderBlocksAsText(_blocks);
            return _fallbackTextProjection != null
                && string.Equals(projectedText, _fallbackTextProjection, StringComparison.Ordinal)
                ? _fallbackText
                : projectedText;
        }
    }
    /// <summary>
    /// Parsed child blocks of the footnote definition.
    /// Paragraph-only footnotes keep this aligned with <see cref="ParagraphBlocks"/>, while richer
    /// producers can preserve headings, code blocks, and other block types here.
    /// </summary>
    public IReadOnlyList<IMarkdownBlock> Blocks => _blocks;
    /// <summary>
    /// Structured child blocks that form the canonical footnote body.
    /// This is the AST-style alias for <see cref="Blocks"/> used by child-container consumers.
    /// </summary>
    public IReadOnlyList<IMarkdownBlock> ChildBlocks => Blocks;
    /// <summary>
    /// Parsed paragraphs of the footnote definition, derived from <see cref="Blocks"/>.
    /// When empty, renderers may fall back to parsing <see cref="Text"/> as a single inline sequence.
    /// </summary>
    public IReadOnlyList<InlineSequence> Paragraphs => CreateParagraphInlines(ParagraphBlocks);
    /// <summary>
    /// Parsed paragraph blocks of the footnote definition, derived from <see cref="Blocks"/>.
    /// This exposes footnote content as owned block children for AST-style consumers.
    /// </summary>
    public IReadOnlyList<ParagraphBlock> ParagraphBlocks => CreateParagraphBlocks(_blocks);
    internal IReadOnlyList<MarkdownSyntaxNode>? SyntaxChildren { get; }
    string IFootnoteSectionMarkdownBlock.FootnoteLabel => Label;

    /// <summary>Create a new footnote definition.</summary>
    /// <param name="label">Identifier used by references.</param>
    /// <param name="text">Definition text.</param>
    public FootnoteDefinitionBlock(string label, string text) {
        Label = label ?? string.Empty;
        _fallbackText = text ?? string.Empty;
        _blocks = CreatePlainTextBodyBlocks(text);
        _fallbackTextProjection = RenderBlocksAsText(_blocks);
    }

    /// <summary>
    /// Creates a footnote definition with structured body blocks.
    /// Prefer this overload when the footnote body contains lists, code blocks, or other nested markdown structure.
    /// </summary>
    public FootnoteDefinitionBlock(string label, IEnumerable<IMarkdownBlock>? childBlocks)
        : this(label, string.Empty, CopyBlocks(childBlocks), syntaxChildren: null) {
    }

    /// <summary>
    /// Creates a footnote definition with structured body blocks and fallback text for empty-block scenarios.
    /// </summary>
    public FootnoteDefinitionBlock(string label, string text, IEnumerable<IMarkdownBlock>? childBlocks)
        : this(label, text, CopyBlocks(childBlocks), syntaxChildren: null) {
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
            CreateBlockList(CreateParagraphBlocks(paragraphs)),
            syntaxChildren) {
    }

    internal FootnoteDefinitionBlock(string label, string text, IReadOnlyList<ParagraphBlock> paragraphBlocks, IReadOnlyList<MarkdownSyntaxNode>? syntaxChildren)
        : this(
            label,
            text,
            CreateBlockList(paragraphBlocks),
            syntaxChildren) {
    }

    internal FootnoteDefinitionBlock(string label, string text, IReadOnlyList<IMarkdownBlock>? blocks, IReadOnlyList<MarkdownSyntaxNode>? syntaxChildren)
        : this(label, text, blocks, fallbackTextProjection: null, syntaxChildren) {
    }

    private FootnoteDefinitionBlock(string label, string text, IReadOnlyList<IMarkdownBlock>? blocks, string? fallbackTextProjection, IReadOnlyList<MarkdownSyntaxNode>? syntaxChildren) {
        Label = label ?? string.Empty;
        _fallbackText = text ?? string.Empty;
        _blocks = CopyBlocks(blocks);
        _fallbackTextProjection = fallbackTextProjection;
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

            mixed.Append(MarkdownBlockRenderDispatcher.RenderHtml(block));
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

                sb.Append(MarkdownBlockRenderDispatcher.RenderHtml(block));
            }

            sb.Append("<p><a class=\"footnote-backref\" href=\"#fnref:").Append(encLabel).Append("\" aria-label=\"Back to reference\">&#8617;</a></p>");
        }

        sb.Append("</li>");
        return sb.ToString();
    }

    IReadOnlyList<IMarkdownBlock> IChildMarkdownBlockContainer.ChildBlocks => ChildBlocks;
    IReadOnlyList<MarkdownSyntaxNode>? ISyntaxChildrenMarkdownBlock.ProvidedSyntaxChildren => SyntaxChildren;

    IReadOnlyList<MarkdownSyntaxNode> IOwnedSyntaxChildrenMarkdownBlock.BuildOwnedSyntaxChildren() {
        return MarkdownBlockSyntaxBuilder.BuildCanonicalChildSyntaxNodes(SyntaxChildren, Blocks);
    }

    MarkdownSyntaxNode ISyntaxMarkdownBlock.BuildSyntaxNode(MarkdownSourceSpan? span) {
        if (!LabelSourceSpan.HasValue || (span.HasValue && !span.Value.Contains(LabelSourceSpan.Value))) {
            LabelSourceSpan = GetFootnoteLabelSpan(span);
        }

        if (!OpeningMarkerSourceSpan.HasValue || (span.HasValue && !span.Value.Contains(OpeningMarkerSourceSpan.Value))) {
            OpeningMarkerSourceSpan = GetFootnoteOpeningMarkerSpan(span);
        }

        if (!SeparatorMarkerSourceSpan.HasValue || (span.HasValue && !span.Value.Contains(SeparatorMarkerSourceSpan.Value))) {
            SeparatorMarkerSourceSpan = GetFootnoteSeparatorMarkerSpan(LabelSourceSpan);
        }

        var children = new List<MarkdownSyntaxNode>();
        if (OpeningMarkerSourceSpan.HasValue && SeparatorMarkerSourceSpan.HasValue) {
            children.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.FootnoteOpeningMarker,
                OpeningMarkerSourceSpan,
                "[^"));
        }

        children.Add(new MarkdownSyntaxNode(
            MarkdownSyntaxKind.FootnoteLabel,
            LabelSourceSpan,
            Label));

        if (OpeningMarkerSourceSpan.HasValue && SeparatorMarkerSourceSpan.HasValue) {
            children.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.FootnoteSeparatorMarker,
                SeparatorMarkerSourceSpan,
                "]:"));
        }

        var bodyChildren = ((IOwnedSyntaxChildrenMarkdownBlock)this).BuildOwnedSyntaxChildren();
        for (int i = 0; i < bodyChildren.Count; i++) {
            children.Add(bodyChildren[i]);
        }

        return new MarkdownSyntaxNode(
            MarkdownSyntaxKind.FootnoteDefinition,
            span,
            Label,
            children,
            this);
    }

    private MarkdownSourceSpan? GetFootnoteLabelSpan(MarkdownSourceSpan? footnoteSpan) {
        if (!footnoteSpan.HasValue || string.IsNullOrEmpty(Label)) {
            return null;
        }

        var startColumn = footnoteSpan.Value.StartColumn;
        if (!startColumn.HasValue) {
            return null;
        }

        return new MarkdownSourceSpan(
            footnoteSpan.Value.StartLine,
            startColumn.Value + 2,
            footnoteSpan.Value.StartLine,
            startColumn.Value + 1 + Label.Length);
    }

    private static MarkdownSourceSpan? GetFootnoteOpeningMarkerSpan(MarkdownSourceSpan? footnoteSpan) {
        if (!footnoteSpan.HasValue || !footnoteSpan.Value.StartColumn.HasValue) {
            return null;
        }

        return new MarkdownSourceSpan(
            footnoteSpan.Value.StartLine,
            footnoteSpan.Value.StartColumn.Value,
            footnoteSpan.Value.StartLine,
            footnoteSpan.Value.StartColumn.Value + 1);
    }

    private static MarkdownSourceSpan? GetFootnoteSeparatorMarkerSpan(MarkdownSourceSpan? labelSpan) {
        if (!labelSpan.HasValue || !labelSpan.Value.EndColumn.HasValue) {
            return null;
        }

        return new MarkdownSourceSpan(
            labelSpan.Value.EndLine,
            labelSpan.Value.EndColumn.Value + 1,
            labelSpan.Value.EndLine,
            labelSpan.Value.EndColumn.Value + 2);
    }

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

    private static IReadOnlyList<IMarkdownBlock> CreatePlainTextBodyBlocks(string? text) {
        if (string.IsNullOrEmpty(text)) {
            return Array.Empty<IMarkdownBlock>();
        }

        return new IMarkdownBlock[] { new ParagraphBlock(MarkdownReader.ParseInlineText(text)) };
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

    private static IReadOnlyList<IMarkdownBlock> CopyBlocks(IEnumerable<IMarkdownBlock>? blocks) {
        if (blocks is IReadOnlyList<IMarkdownBlock> readOnlyBlocks) {
            return CopyBlocks(readOnlyBlocks);
        }

        if (blocks == null) {
            return Array.Empty<IMarkdownBlock>();
        }

        return blocks.Where(block => block != null).ToArray();
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
