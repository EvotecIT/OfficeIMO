namespace OfficeIMO.Markdown;

/// <summary>
/// Sequence of inline nodes used in paragraphs and list items.
/// </summary>
public sealed class InlineSequence : MarkdownInline, IRenderableMarkdownInline, IPlainTextMarkdownInline {
    private readonly List<IMarkdownInline> _inlines = new List<IMarkdownInline>();
    private readonly IReadOnlyList<object> _itemsView;

    /// <summary>Creates an empty inline sequence.</summary>
    public InlineSequence() {
        _itemsView = new InlineObjectReadOnlyList(_inlines);
    }

    // When composing via the fluent/builder APIs, auto-spacing between adjacent inline nodes is convenient.
    // When parsing Markdown source, spacing is already present in TextRun nodes, so auto-spacing would double spaces.
    internal bool AutoSpacing { get; set; } = true;
    /// <summary>Exposes the inline nodes for safe iteration.</summary>
    public IReadOnlyList<IMarkdownInline> Nodes => _inlines;
    /// <summary>Legacy object-typed inline view retained for compatibility.</summary>
    public IReadOnlyList<object> Items => _itemsView;
    /// <summary>Adds plain text.</summary>
    public InlineSequence Text(string text) { _inlines.Add(new TextRun(text)); return this; }
    /// <summary>Adds a hyperlink.</summary>
    public InlineSequence Link(string text, string url, string? title = null, string? linkTarget = null, string? linkRel = null) { _inlines.Add(new LinkInline(text, url, title, linkTarget, linkRel)); return this; }
    /// <summary>Adds bold text.</summary>
    public InlineSequence Bold(string text) { _inlines.Add(new BoldInline(text)); return this; }
    /// <summary>Adds bold+italic text.</summary>
    public InlineSequence BoldItalic(string text) { _inlines.Add(new BoldItalicInline(text)); return this; }
    /// <summary>Adds italic text.</summary>
    public InlineSequence Italic(string text) { _inlines.Add(new ItalicInline(text)); return this; }
    /// <summary>Adds inline code.</summary>
    public InlineSequence Code(string text) { _inlines.Add(new CodeSpanInline(text)); return this; }
    /// <summary>Adds a footnote reference (e.g., [^id]).</summary>
    public InlineSequence FootnoteRef(string label) { _inlines.Add(new FootnoteRefInline(label)); return this; }
    /// <summary>Adds strikethrough text.</summary>
    public InlineSequence Strike(string text) { _inlines.Add(new StrikethroughInline(text)); return this; }
    /// <summary>Adds highlighted text rendered as <c>==text==</c>.</summary>
    public InlineSequence Highlight(string text) { _inlines.Add(new HighlightInline(text)); return this; }
    /// <summary>Adds underlined text (HTML-only in Markdown).</summary>
    public InlineSequence Underline(string text) { _inlines.Add(new UnderlineInline(text)); return this; }
    /// <summary>Adds superscript text rendered via inline HTML.</summary>
    public InlineSequence Superscript(string text) { _inlines.Add(new HtmlTagSequenceInline("sup", new InlineSequence().Text(text))); return this; }
    /// <summary>Adds subscript text rendered via inline HTML.</summary>
    public InlineSequence Subscript(string text) { _inlines.Add(new HtmlTagSequenceInline("sub", new InlineSequence().Text(text))); return this; }
    /// <summary>Adds inserted text rendered via inline HTML.</summary>
    public InlineSequence Inserted(string text) { _inlines.Add(new HtmlTagSequenceInline("ins", new InlineSequence().Text(text))); return this; }
    /// <summary>Adds quoted text rendered via inline HTML.</summary>
    public InlineSequence Quote(string text) { _inlines.Add(new HtmlTagSequenceInline("q", new InlineSequence().Text(text))); return this; }
    /// <summary>Adds a linked image (useful for badges).</summary>
    public InlineSequence ImageLink(string alt, string imageUrl, string linkUrl, string? title = null, string? linkTitle = null) { _inlines.Add(new ImageLinkInline(alt, imageUrl, linkUrl, title, linkTitle)); return this; }
    /// <summary>Adds a standalone inline image.</summary>
    public InlineSequence Image(string alt, string src, string? title = null) { _inlines.Add(new ImageInline(alt, src, title)); return this; }
    /// <summary>Adds a hard line break.</summary>
    public InlineSequence HardBreak() { _inlines.Add(new HardBreakInline()); return this; }

    // Internal escape hatch for the reader to attach richer inline nodes without expanding the public fluent API.
    internal InlineSequence AddRaw(IMarkdownInline node) { if (node != null) _inlines.Add(node); return this; }

    internal void ReplaceItems(IEnumerable<IMarkdownInline> nodes) {
        _inlines.Clear();
        if (nodes == null) {
            return;
        }

        foreach (var node in nodes) {
            if (node != null) {
                _inlines.Add(node);
            }
        }
    }

    internal string RenderMarkdown() {
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < _inlines.Count; i++) {
            if (AutoSpacing && i > 0) {
                var prev = _inlines[i - 1];
                var cur = _inlines[i];
                if (prev is not HardBreakInline && cur is not HardBreakInline) sb.Append(' ');
            }
            sb.Append(GetRenderable(_inlines[i]).RenderMarkdown());
        }
        return sb.ToString();
    }

    internal string RenderHtml() {
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < _inlines.Count; i++) {
            if (AutoSpacing && i > 0) {
                var prev = _inlines[i - 1];
                var cur = _inlines[i];
                if (prev is not HardBreakInline && cur is not HardBreakInline) sb.Append(' ');
            }
            sb.Append(GetRenderable(_inlines[i]).RenderHtml());
        }
        return sb.ToString();
    }

    string IRenderableMarkdownInline.RenderMarkdown() => RenderMarkdown();
    string IRenderableMarkdownInline.RenderHtml() => RenderHtml();
    void IPlainTextMarkdownInline.AppendPlainText(System.Text.StringBuilder sb) => InlinePlainText.AppendPlainText(sb, this);

    private static IRenderableMarkdownInline GetRenderable(IMarkdownInline node) {
        return node as IRenderableMarkdownInline
            ?? throw new InvalidOperationException($"Inline node of type '{node.GetType().FullName}' does not implement {nameof(IRenderableMarkdownInline)}.");
    }
}

internal sealed class InlineObjectReadOnlyList : IReadOnlyList<object> {
    private readonly IReadOnlyList<IMarkdownInline> _nodes;

    public InlineObjectReadOnlyList(IReadOnlyList<IMarkdownInline> nodes) {
        _nodes = nodes ?? throw new ArgumentNullException(nameof(nodes));
    }

    public int Count => _nodes.Count;

    public object this[int index] => _nodes[index];

    public IEnumerator<object> GetEnumerator() {
        for (int i = 0; i < _nodes.Count; i++) {
            yield return _nodes[i];
        }
    }

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
}
