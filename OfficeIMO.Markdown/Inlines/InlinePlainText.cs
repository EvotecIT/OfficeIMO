namespace OfficeIMO.Markdown;

internal static class InlinePlainText {
    public static string Extract(InlineSequence? sequence) {
        if (sequence == null || sequence.Items.Count == 0) {
            return string.Empty;
        }

        var sb = new System.Text.StringBuilder();
        Append(sb, sequence);
        return sb.ToString();
    }

    private static void Append(System.Text.StringBuilder sb, InlineSequence sequence) {
        foreach (var node in sequence.Items) {
            if (node is TextRun t) sb.Append(t.Text);
            else if (node is CodeSpanInline cs) sb.Append(cs.Text);
            else if (node is ItalicSequenceInline it) Append(sb, it.Inlines);
            else if (node is BoldSequenceInline b) Append(sb, b.Inlines);
            else if (node is BoldItalicSequenceInline bi) Append(sb, bi.Inlines);
            else if (node is StrikethroughSequenceInline st) Append(sb, st.Inlines);
            else if (node is HighlightSequenceInline hi) Append(sb, hi.Inlines);
            else if (node is HardBreakInline) sb.Append(' ');
            else if (node is UnderlineInline u) sb.Append(u.Text);
            else if (node is HighlightInline h) sb.Append(h.Text);
            else if (node is FootnoteRefInline fn) sb.Append(fn.Label);
            else if (node is LinkInline l) sb.Append(l.Text);
            else if (node is ImageInline im) sb.Append(im.Alt);
            else if (node is ImageLinkInline il) sb.Append(il.Alt);
            else if (node is ItalicInline italic) sb.Append(italic.Text);
            else if (node is BoldInline bold) sb.Append(bold.Text);
            else if (node is BoldItalicInline boldItalic) sb.Append(boldItalic.Text);
            else if (node is StrikethroughInline strike) sb.Append(strike.Text);
        }
    }
}
