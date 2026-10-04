namespace OfficeIMO.AsciiDoc.Markdown;

internal static class MarkdownInlineToAsciiDocConverter {
    internal static string Convert(
        InlineSequence source,
        List<MarkdownAsciiDocConversionDiagnostic> diagnostics,
        MarkdownObject owner) {
        var output = new System.Text.StringBuilder();
        for (int index = 0; index < source.Nodes.Count; index++) {
            IMarkdownInline inline = source.Nodes[index];
            switch (inline) {
                case MarkdownTextRun text: output.Append(EscapeText(text.Text)); break;
                case BoldInline bold: AppendFormatted(output, "*", EscapeText(bold.Text), source, index); break;
                case BoldSequenceInline bold: AppendFormatted(output, "*", Convert(bold.Inlines, diagnostics, owner), source, index); break;
                case BoldItalicInline boldItalic: AppendFormatted(output, "*_", EscapeText(boldItalic.Text), source, index); break;
                case BoldItalicSequenceInline boldItalic: AppendFormatted(output, "*_", Convert(boldItalic.Inlines, diagnostics, owner), source, index); break;
                case ItalicInline italic: AppendFormatted(output, "_", EscapeText(italic.Text), source, index); break;
                case ItalicSequenceInline italic: AppendFormatted(output, "_", Convert(italic.Inlines, diagnostics, owner), source, index); break;
                case CodeSpanInline code: AppendFormatted(output, "`", EscapeText(code.Text), source, index); break;
                case HighlightInline highlight: AppendFormatted(output, "#", EscapeText(highlight.Text), source, index); break;
                case HighlightSequenceInline highlight: AppendFormatted(output, "#", Convert(highlight.Inlines, diagnostics, owner), source, index); break;
                case SuperscriptInline superscript: output.Append('^').Append(EscapeText(superscript.Text)).Append('^'); break;
                case SuperscriptSequenceInline superscript: output.Append('^').Append(Convert(superscript.Inlines, diagnostics, owner)).Append('^'); break;
                case SubscriptInline subscript: output.Append('~').Append(EscapeText(subscript.Text)).Append('~'); break;
                case SubscriptSequenceInline subscript: output.Append('~').Append(Convert(subscript.Inlines, diagnostics, owner)).Append('~'); break;
                case StrikethroughInline strike: output.Append("[line-through]#").Append(EscapeText(strike.Text)).Append('#'); break;
                case StrikethroughSequenceInline strike: output.Append("[line-through]#").Append(Convert(strike.Inlines, diagnostics, owner)).Append('#'); break;
                case InsertedInline inserted: output.Append("[.inserted]#").Append(EscapeText(inserted.Text)).Append('#'); break;
                case InsertedSequenceInline inserted: output.Append("[.inserted]#").Append(Convert(inserted.Inlines, diagnostics, owner)).Append('#'); break;
                case LinkInline link: AddLink(output, link, diagnostics, owner); break;
                case ImageInline image: output.Append("image:").Append(EscapeTarget(image.Src)).Append('[').Append(EscapeAttribute(image.Alt)).Append(']'); break;
                case FootnoteRefInline footnote: output.Append("footnote:").Append(EscapeTarget(footnote.Label)).Append("[]"); break;
                case HardBreakInline: output.Append(" +\n"); break;
                case SoftBreakInline: output.Append('\n'); break;
                case HtmlRawInline html:
                    output.Append("pass:[").Append(html.Html.Replace("]", "\\]")).Append(']');
                    Report(diagnostics, owner, "MDADOC103", "raw-html-inline", "Raw HTML retained in an AsciiDoc pass macro.");
                    break;
                default:
                    output.Append("pass:[").Append(RenderInlineFallback(inline).Replace("]", "\\]")).Append(']');
                    Report(diagnostics, owner, "MDADOC109", inline.GetType().Name, "Unsupported Markdown inline retained in a pass macro.");
                    break;
            }
        }
        return output.ToString();
    }

    private static void AddLink(
        System.Text.StringBuilder output,
        LinkInline link,
        List<MarkdownAsciiDocConversionDiagnostic> diagnostics,
        MarkdownObject owner) {
        string label = link.LabelInlines == null ? EscapeAttribute(link.Text) : Convert(link.LabelInlines, diagnostics, owner);
        if (link.Url.StartsWith("#", StringComparison.Ordinal) && link.Url.Length > 1) {
            output.Append("<<").Append(EscapeTarget(link.Url.Substring(1))).Append(',').Append(label).Append(">>");
        } else {
            output.Append("link:").Append(EscapeTarget(link.Url)).Append('[').Append(label);
            if (!string.IsNullOrEmpty(link.Title)) output.Append(",\"").Append(EscapeAttribute(link.Title!)).Append('"');
            output.Append(']');
        }
    }

    private static string RenderInlineFallback(IMarkdownInline inline) {
        var sequence = new InlineSequence { AutoSpacing = false };
        sequence.AddRaw(inline);
        return MarkdownDoc.Create().Add(new ParagraphBlock(sequence)).ToMarkdown();
    }

    private static string EscapeText(string value) {
        var output = new System.Text.StringBuilder(value.Length);
        for (int index = 0; index < value.Length; index++) {
            char current = value[index];
            if ("\\*_`#~^{}[]<>".IndexOf(current) >= 0) output.Append('\\');
            output.Append(current);
        }
        return output.ToString();
    }

    private static void AppendFormatted(System.Text.StringBuilder output, string markers, string content, InlineSequence source, int index) {
        char? before = output.Length == 0 ? null : (char?)output[output.Length - 1];
        char? after = null;
        if (index + 1 < source.Nodes.Count) {
            IMarkdownInline next = source.Nodes[index + 1];
            after = next is MarkdownTextRun text && text.Text.Length > 0 ? text.Text[0]
                : next is SoftBreakInline || next is HardBreakInline ? '\n' : '_';
        }
        bool unconstrained = !AsciiDocInlineFormattingRules.CanOpen(before, content.Length == 0 ? null : (char?)content[0]) ||
            !AsciiDocInlineFormattingRules.CanClose(content.Length == 0 ? null : (char?)content[content.Length - 1], after);
        string opening = unconstrained ? string.Concat(markers.Select(marker => new string(marker, 2))) : markers;
        string closing = new string(opening.Reverse().ToArray());
        output.Append(opening).Append(content).Append(closing);
    }

    private static string EscapeTarget(string value) => value.Replace("[", "\\[").Replace("]", "\\]");

    private static string EscapeAttribute(string value) => value.Replace("\\", "\\\\").Replace("]", "\\]").Replace(",", "\\,").Replace("\"", "\\\"");

    private static void Report(
        List<MarkdownAsciiDocConversionDiagnostic> diagnostics,
        MarkdownObject owner,
        string code,
        string feature,
        string message) {
        diagnostics.Add(new MarkdownAsciiDocConversionDiagnostic(
            code,
            AsciiDocMarkdownDiagnosticSeverity.Warning,
            AsciiDocMarkdownConversionOutcome.Simplified,
            feature,
            message,
            owner.SourceSpan));
    }
}
