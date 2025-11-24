namespace OfficeIMO.Markdown;

/// <summary>
/// Collapsible disclosure block with an optional summary and nested content.
/// </summary>
public sealed class DetailsBlock : IMarkdownBlock {
    /// <summary>Optional summary displayed in the disclosure header.</summary>
    public SummaryBlock? Summary { get; set; }

    /// <summary>Nested blocks rendered inside the details body.</summary>
    public System.Collections.Generic.List<IMarkdownBlock> Children { get; } = new System.Collections.Generic.List<IMarkdownBlock>();

    /// <summary>Whether to emit a blank line between the summary and the first child block.</summary>
    public bool InsertBlankLineAfterSummary { get; set; } = true;

    /// <summary>Whether to emit a blank line before the closing tag.</summary>
    public bool InsertBlankLineBeforeClosing { get; set; }

    /// <summary>Whether the details element is initially expanded.</summary>
    public bool Open { get; set; }

    /// <summary>Creates an empty details block.</summary>
    public DetailsBlock() {
    }

    /// <summary>Creates a details block with a summary and children.</summary>
    public DetailsBlock(SummaryBlock? summary, System.Collections.Generic.IEnumerable<IMarkdownBlock>? children = null, bool open = false) {
        Summary = summary;
        Open = open;
        if (children != null) Children.AddRange(children);
    }

    string IMarkdownBlock.RenderMarkdown() => Render(renderHtmlChildren: false);
    string IMarkdownBlock.RenderHtml() => Render(renderHtmlChildren: true);

    private string Render(bool renderHtmlChildren) {
        var sb = new System.Text.StringBuilder();
        const string NewLine = "\n";
        sb.Append("<details");
        if (Open) sb.Append(" open");
        sb.Append('>');

        if (Summary != null) {
            sb.Append(NewLine);
            sb.Append(((IMarkdownBlock)Summary).RenderHtml());
        }

        if (Children.Count > 0) {
            sb.Append(NewLine);
            for (int i = 0; i < Children.Count; i++) {
                if (i == 0) {
                    if (Summary != null && InsertBlankLineAfterSummary) sb.Append(NewLine);
                } else {
                    sb.Append(NewLine).Append(NewLine);
                }
                var rendered = renderHtmlChildren ? Children[i].RenderHtml() : Children[i].RenderMarkdown();
                sb.Append(rendered);
            }
        }

        sb.Append(NewLine);
        if (InsertBlankLineBeforeClosing && (Children.Count > 0 || Summary != null)) sb.Append(NewLine);
        sb.Append("</details>");
        return sb.ToString();
    }
}

/// <summary>
/// Summary header for a <see cref="DetailsBlock"/>.
/// </summary>
public sealed class SummaryBlock : IMarkdownBlock {
    /// <summary>Inline content inside the &lt;summary&gt; element.</summary>
    public InlineSequence Inlines { get; }

    /// <summary>Create a summary block from an inline sequence.</summary>
    public SummaryBlock(InlineSequence inlines) {
        Inlines = inlines ?? new InlineSequence();
    }

    /// <summary>Create a summary block with plain text.</summary>
    public SummaryBlock(string? text) {
        Inlines = new InlineSequence().Text(text ?? string.Empty);
    }

    string IMarkdownBlock.RenderMarkdown() => $"<summary>{Inlines.RenderHtml()}</summary>";
    string IMarkdownBlock.RenderHtml() => $"<summary>{Inlines.RenderHtml()}</summary>";
}
