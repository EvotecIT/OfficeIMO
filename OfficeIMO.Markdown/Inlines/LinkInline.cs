namespace OfficeIMO.Markdown;

/// <summary>
/// Hyperlink inline.
/// </summary>
public sealed class LinkInline {
    /// <summary>Link text.</summary>
    public string Text { get; }
    /// <summary>Destination URL.</summary>
    public string Url { get; }
    /// <summary>Optional title shown as a tooltip in HTML.</summary>
    public string? Title { get; }

    // Optional richer label representation (produced by the reader). When present, RenderHtml/RenderMarkdown
    // uses it instead of the plain Text property.
    internal InlineSequence? LabelInlines { get; }
    /// <summary>Creates a hyperlink inline.</summary>
    public LinkInline(string text, string url, string? title) { Text = text ?? string.Empty; Url = url ?? string.Empty; Title = title; }

    internal LinkInline(InlineSequence label, string url, string? title)
        : this(ExtractPlainText(label), url, title) {
        LabelInlines = label;
    }

    private static string ExtractPlainText(InlineSequence label) {
        if (label == null) return string.Empty;
        if (label.Items == null || label.Items.Count == 0) return string.Empty;

        var sb = new System.Text.StringBuilder();
        AppendPlainText(sb, label);
        return sb.ToString();
    }

    private static void AppendPlainText(System.Text.StringBuilder sb, InlineSequence seq) {
        if (sb == null || seq == null) return;
        foreach (var node in seq.Items) {
            if (node is TextRun t) sb.Append(t.Text);
            else if (node is CodeSpanInline cs) sb.Append(cs.Text);
            else if (node is ItalicSequenceInline it) AppendPlainText(sb, it.Inlines);
            else if (node is BoldSequenceInline b) AppendPlainText(sb, b.Inlines);
            else if (node is BoldItalicSequenceInline bi) AppendPlainText(sb, bi.Inlines);
            else if (node is StrikethroughSequenceInline st) AppendPlainText(sb, st.Inlines);
            else if (node is HardBreakInline) sb.Append(' ');
            else if (node is UnderlineInline u) sb.Append(u.Text);
            else if (node is FootnoteRefInline fn) sb.Append(fn.Label);
            // Links/images inside link labels are not expected when produced by our reader (they're disabled),
            // but if they occur, keep their plain text parts.
            else if (node is LinkInline l) sb.Append(l.Text);
            else if (node is ImageInline im) sb.Append(im.Alt);
            else if (node is ImageLinkInline il) sb.Append(il.Alt);
        }
    }
    internal string RenderMarkdown() {
        string title = MarkdownEscaper.FormatOptionalTitle(Title);
        if (LabelInlines != null) {
            return $"[{LabelInlines.RenderMarkdown()}]({MarkdownEscaper.EscapeLinkUrl(Url)}{title})";
        }
        return $"[{MarkdownEscaper.EscapeLinkText(Text)}]({MarkdownEscaper.EscapeLinkUrl(Url)}{title})";
    }
    internal string RenderHtml() {
        string title = string.IsNullOrEmpty(Title) ? string.Empty : $" title=\"{System.Net.WebUtility.HtmlEncode(Title!)}\"";
        if (LabelInlines != null) {
            return $"<a href=\"{System.Net.WebUtility.HtmlEncode(Url)}\"{title}>{LabelInlines.RenderHtml()}</a>";
        }
        return $"<a href=\"{System.Net.WebUtility.HtmlEncode(Url)}\"{title}>{System.Net.WebUtility.HtmlEncode(Text)}</a>";
    }
}
