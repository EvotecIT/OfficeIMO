using System;
using System.Text;

namespace OfficeIMO.Markdown;

/// <summary>
/// Typed table cell containing one or more markdown blocks.
/// </summary>
public sealed class TableCell : MarkdownObject, IChildMarkdownBlockContainer, ISyntaxChildrenMarkdownBlock, IOwnedSyntaxChildrenMarkdownBlock {
    private int _columnSpan = 1;
    private int _rowSpan = 1;

    /// <summary>Structured cell content.</summary>
    public List<IMarkdownBlock> Blocks { get; } = new List<IMarkdownBlock>();
    /// <summary>Structured child blocks owned by this table cell.</summary>
    public IReadOnlyList<IMarkdownBlock> ChildBlocks => Blocks;
    /// <summary>Owned syntax nodes for the structured cell body.</summary>
    internal IReadOnlyList<MarkdownSyntaxNode>? SyntaxChildren { get; set; }
    /// <summary>Whether this cell belongs to the header row.</summary>
    public bool IsHeader { get; internal set; }
    /// <summary>Zero-based data-row index for body cells; <c>-1</c> for header cells.</summary>
    public int RowIndex { get; internal set; } = -1;
    /// <summary>Zero-based column index within the row.</summary>
    public int ColumnIndex { get; internal set; } = -1;
    /// <summary>Optional cell-level horizontal alignment override.</summary>
    public ColumnAlignment Alignment { get; set; } = ColumnAlignment.None;
    /// <summary>Optional CSS-compatible cell background color token.</summary>
    public string? BackgroundColor { get; set; }
    /// <summary>Optional CSS-compatible cell text color token.</summary>
    public string? TextColor { get; set; }
    /// <summary>Whether cell text should be rendered with bold emphasis.</summary>
    public bool Bold { get; set; }
    /// <summary>Whether cell text should be rendered with italic emphasis.</summary>
    public bool Italic { get; set; }
    /// <summary>Whether cell text should be rendered with underline decoration.</summary>
    public bool Underline { get; set; }
    /// <summary>Whether cell text should be rendered with strikethrough decoration.</summary>
    public bool Strikethrough { get; set; }
    /// <summary>Number of logical table columns covered by this cell.</summary>
    public int ColumnSpan {
        get => _columnSpan;
        set {
            if (value < 1) {
                throw new ArgumentOutOfRangeException(nameof(value), "Table cell column span must be at least 1.");
            }

            _columnSpan = value;
        }
    }
    /// <summary>Number of logical table rows covered by this cell.</summary>
    public int RowSpan {
        get => _rowSpan;
        set {
            if (value < 1) {
                throw new ArgumentOutOfRangeException(nameof(value), "Table cell row span must be at least 1.");
            }

            _rowSpan = value;
        }
    }
    /// <summary>Creates a typed table cell.</summary>
    public TableCell(IEnumerable<IMarkdownBlock>? blocks = null) {
        if (blocks == null) {
            return;
        }

        foreach (var block in blocks) {
            if (block != null) {
                Blocks.Add(block);
            }
        }
    }

    /// <summary>Markdown representation of the full cell body.</summary>
    public string Markdown => RenderMarkdown();

    internal string RenderMarkdown() {
        if (Blocks.Count == 0) {
            return string.Empty;
        }

        if (Blocks.Count == 1 && Blocks[0] is ParagraphBlock paragraph) {
            return paragraph.Inlines.RenderMarkdown();
        }

        var sb = new StringBuilder();
        for (int i = 0; i < Blocks.Count; i++) {
            var rendered = MarkdownBlockRenderDispatcher.RenderMarkdown(Blocks[i]);
            if (string.IsNullOrEmpty(rendered)) {
                continue;
            }

            if (sb.Length > 0) {
                sb.Append("\n\n");
            }

            sb.Append(rendered);
        }

        return sb.ToString();
    }

    internal string RenderHtml() {
        if (Blocks.Count == 0) {
            return string.Empty;
        }

        if (Blocks.Count == 1 && Blocks[0] is ParagraphBlock paragraph) {
            var rendered = paragraph.Inlines.RenderHtml();
            return rendered.Contains('\n') ? rendered.Replace("\n", "<br/>") : rendered;
        }

        var sb = new StringBuilder();
        for (int i = 0; i < Blocks.Count; i++) {
            sb.Append(MarkdownBlockRenderDispatcher.RenderHtml(Blocks[i]));
        }

        return sb.ToString();
    }

    IReadOnlyList<MarkdownSyntaxNode>? ISyntaxChildrenMarkdownBlock.ProvidedSyntaxChildren => SyntaxChildren;

    IReadOnlyList<MarkdownSyntaxNode> IOwnedSyntaxChildrenMarkdownBlock.BuildOwnedSyntaxChildren() {
        if (SyntaxChildren != null
            && SyntaxChildren.Count > 0
            && MarkdownBlockSyntaxBuilder.ChildSyntaxNodesMatchBlocks(SyntaxChildren, Blocks)) {
            return SyntaxChildren;
        }

        return MarkdownBlockSyntaxBuilder.BuildChildSyntaxNodes(ChildBlocks);
    }
}
