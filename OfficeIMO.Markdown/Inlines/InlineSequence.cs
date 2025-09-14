using System.Collections.Generic;
using System.Text;

namespace OfficeIMO.Markdown;

/// <summary>
/// Sequence of inline nodes used in paragraphs and list items.
/// </summary>
public sealed class InlineSequence : IMarkdownInline {
    private readonly List<object> _inlines = new List<object>();
    /// <summary>Exposes the inline nodes for safe iteration.</summary>
    public IReadOnlyList<object> Items => _inlines;
    /// <summary>Adds plain text.</summary>
    public InlineSequence Text(string text) { _inlines.Add(new TextRun(text)); return this; }
    /// <summary>Adds a hyperlink.</summary>
    public InlineSequence Link(string text, string url, string? title = null) { _inlines.Add(new LinkInline(text, url, title)); return this; }
    /// <summary>Adds bold text.</summary>
    public InlineSequence Bold(string text) { _inlines.Add(new BoldInline(text)); return this; }
    public InlineSequence BoldItalic(string text) { _inlines.Add(new BoldItalicInline(text)); return this; }
    /// <summary>Adds italic text.</summary>
    public InlineSequence Italic(string text) { _inlines.Add(new ItalicInline(text)); return this; }
    /// <summary>Adds inline code.</summary>
    public InlineSequence Code(string text) { _inlines.Add(new CodeSpanInline(text)); return this; }
    public InlineSequence FootnoteRef(string label) { _inlines.Add(new FootnoteRefInline(label)); return this; }
    /// <summary>Adds strikethrough text.</summary>
    public InlineSequence Strike(string text) { _inlines.Add(new StrikethroughInline(text)); return this; }
    /// <summary>Adds underlined text (HTML-only in Markdown).</summary>
    public InlineSequence Underline(string text) { _inlines.Add(new UnderlineInline(text)); return this; }
    /// <summary>Adds a linked image (useful for badges).</summary>
    public InlineSequence ImageLink(string alt, string imageUrl, string linkUrl, string? title = null) { _inlines.Add(new ImageLinkInline(alt, imageUrl, linkUrl, title)); return this; }
    /// <summary>Adds a hard line break.</summary>
    public InlineSequence HardBreak() { _inlines.Add(new HardBreakInline()); return this; }

    internal string RenderMarkdown() {
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < _inlines.Count; i++) {
            if (i > 0) {
                var prev = _inlines[i - 1];
                var cur = _inlines[i];
                if (prev is not HardBreakInline && cur is not HardBreakInline) sb.Append(' ');
            }
            object node = _inlines[i];
            if (node is TextRun t) sb.Append(t.RenderMarkdown());
            else if (node is LinkInline l) sb.Append(l.RenderMarkdown());
            else if (node is BoldInline b) sb.Append(b.RenderMarkdown());
            else if (node is BoldItalicInline bi) sb.Append(bi.RenderMarkdown());
            else if (node is ItalicInline it) sb.Append(it.RenderMarkdown());
            else if (node is CodeSpanInline cs) sb.Append(cs.RenderMarkdown());
            else if (node is ImageLinkInline il) sb.Append(il.RenderMarkdown());
            else if (node is StrikethroughInline st) sb.Append(st.RenderMarkdown());
            else if (node is UnderlineInline un) sb.Append(un.RenderMarkdown());
            else if (node is FootnoteRefInline fn) sb.Append(fn.RenderMarkdown());
        }
        return sb.ToString();
    }

    internal string RenderHtml() {
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < _inlines.Count; i++) {
            if (i > 0) {
                var prev = _inlines[i - 1];
                var cur = _inlines[i];
                if (prev is not HardBreakInline && cur is not HardBreakInline) sb.Append(' ');
            }
            object node = _inlines[i];
            if (node is TextRun t) sb.Append(t.RenderHtml());
            else if (node is LinkInline l) sb.Append(l.RenderHtml());
            else if (node is BoldInline b) sb.Append(b.RenderHtml());
            else if (node is BoldItalicInline bi) sb.Append(bi.RenderHtml());
            else if (node is ItalicInline it) sb.Append(it.RenderHtml());
            else if (node is CodeSpanInline cs) sb.Append(cs.RenderHtml());
            else if (node is ImageLinkInline il) sb.Append(il.RenderHtml());
            else if (node is StrikethroughInline st) sb.Append(st.RenderHtml());
            else if (node is UnderlineInline un) sb.Append(un.RenderHtml());
            else if (node is FootnoteRefInline fn) sb.Append(fn.RenderHtml());
        }
        return sb.ToString();
    }
}
