namespace OfficeIMO.Markdown;

/// <summary>
/// Sequence of inline nodes used in paragraphs and list items.
/// </summary>
public sealed class InlineSequence : IMarkdownInline {
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
    public InlineSequence Link(string text, string url, string? title = null) { _inlines.Add(new LinkInline(text, url, title)); return this; }
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
    /// <summary>Adds a linked image (useful for badges).</summary>
    public InlineSequence ImageLink(string alt, string imageUrl, string linkUrl, string? title = null) { _inlines.Add(new ImageLinkInline(alt, imageUrl, linkUrl, title)); return this; }
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
            var node = _inlines[i];
            if (node is TextRun t) sb.Append(t.RenderMarkdown());
            else if (node is LinkInline l) sb.Append(l.RenderMarkdown());
            else if (node is BoldInline b) sb.Append(b.RenderMarkdown());
            else if (node is BoldItalicInline bi) sb.Append(bi.RenderMarkdown());
            else if (node is ItalicInline it) sb.Append(it.RenderMarkdown());
            else if (node is CodeSpanInline cs) sb.Append(cs.RenderMarkdown());
            else if (node is ImageLinkInline il) sb.Append(il.RenderMarkdown());
            else if (node is ImageInline im) sb.Append(im.RenderMarkdown());
            else if (node is StrikethroughInline st) sb.Append(st.RenderMarkdown());
            else if (node is HighlightInline hi) sb.Append(hi.RenderMarkdown());
            else if (node is UnderlineInline un) sb.Append(un.RenderMarkdown());
            else if (node is FootnoteRefInline fn) sb.Append(fn.RenderMarkdown());
            else if (node is HardBreakInline hb) sb.Append(hb.RenderMarkdown());
            else if (node is BoldSequenceInline bs) sb.Append(bs.RenderMarkdown());
            else if (node is ItalicSequenceInline es) sb.Append(es.RenderMarkdown());
            else if (node is BoldItalicSequenceInline bis) sb.Append(bis.RenderMarkdown());
            else if (node is StrikethroughSequenceInline sts) sb.Append(sts.RenderMarkdown());
            else if (node is HighlightSequenceInline hs) sb.Append(hs.RenderMarkdown());
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
            var node = _inlines[i];
            if (node is TextRun t) sb.Append(t.RenderHtml());
            else if (node is LinkInline l) sb.Append(l.RenderHtml());
            else if (node is BoldInline b) sb.Append(b.RenderHtml());
            else if (node is BoldItalicInline bi) sb.Append(bi.RenderHtml());
            else if (node is ItalicInline it) sb.Append(it.RenderHtml());
            else if (node is CodeSpanInline cs) sb.Append(cs.RenderHtml());
            else if (node is ImageLinkInline il) sb.Append(il.RenderHtml());
            else if (node is ImageInline im) sb.Append(im.RenderHtml());
            else if (node is StrikethroughInline st) sb.Append(st.RenderHtml());
            else if (node is HighlightInline hi) sb.Append(hi.RenderHtml());
            else if (node is UnderlineInline un) sb.Append(un.RenderHtml());
            else if (node is FootnoteRefInline fn) sb.Append(fn.RenderHtml());
            else if (node is HardBreakInline hb) sb.Append(hb.RenderHtml());
            else if (node is BoldSequenceInline bs) sb.Append(bs.RenderHtml());
            else if (node is ItalicSequenceInline es) sb.Append(es.RenderHtml());
            else if (node is BoldItalicSequenceInline bis) sb.Append(bis.RenderHtml());
            else if (node is StrikethroughSequenceInline sts) sb.Append(sts.RenderHtml());
            else if (node is HighlightSequenceInline hs) sb.Append(hs.RenderHtml());
        }
        return sb.ToString();
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
