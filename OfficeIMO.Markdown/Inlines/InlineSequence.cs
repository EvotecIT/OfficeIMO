using System.Collections.Generic;
using System.Text;

namespace OfficeIMO.Markdown;

public sealed class InlineSequence : IMarkdownInline {
    private readonly List<object> _inlines = new List<object>();
    public InlineSequence Text(string text) { _inlines.Add(new TextRun(text)); return this; }
    public InlineSequence Link(string text, string url, string? title = null) { _inlines.Add(new LinkInline(text, url, title)); return this; }

    internal string RenderMarkdown() {
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < _inlines.Count; i++) {
            if (i > 0) sb.Append(' ');
            sb.Append(_inlines[i] switch {
                TextRun t => t.RenderMarkdown(),
                LinkInline l => l.RenderMarkdown(),
                _ => string.Empty
            });
        }
        return sb.ToString();
    }

    internal string RenderHtml() {
        StringBuilder sb = new StringBuilder();
        foreach (object node in _inlines) {
            sb.Append(node switch {
                TextRun t => t.RenderHtml(),
                LinkInline l => l.RenderHtml(),
                _ => string.Empty
            });
            sb.Append(' ');
        }
        return sb.ToString().TrimEnd();
    }
}

