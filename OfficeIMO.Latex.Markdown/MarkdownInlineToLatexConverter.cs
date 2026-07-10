namespace OfficeIMO.Latex.Markdown;

internal static class MarkdownInlineToLatexConverter {
    internal static string Convert(
        InlineSequence source,
        ConversionState state,
        List<LatexMarkdownConversionDiagnostic> diagnostics,
        MarkdownObject owner) {
        var output = new StringBuilder();
        for (int index = 0; index < source.Nodes.Count; index++) {
            IMarkdownInline node = source.Nodes[index];
            switch (node) {
                case TextRun text: output.Append(EscapeText(text.Text)); break;
                case BoldInline bold: output.Append("\\textbf{").Append(EscapeText(bold.Text)).Append('}'); break;
                case BoldSequenceInline bold: output.Append("\\textbf{").Append(Convert(bold.Inlines, state, diagnostics, owner)).Append('}'); break;
                case BoldItalicInline boldItalic: output.Append("\\textbf{\\emph{").Append(EscapeText(boldItalic.Text)).Append("}}"); break;
                case BoldItalicSequenceInline boldItalic: output.Append("\\textbf{\\emph{").Append(Convert(boldItalic.Inlines, state, diagnostics, owner)).Append("}}"); break;
                case ItalicInline italic: output.Append("\\emph{").Append(EscapeText(italic.Text)).Append('}'); break;
                case ItalicSequenceInline italic: output.Append("\\emph{").Append(Convert(italic.Inlines, state, diagnostics, owner)).Append('}'); break;
                case CodeSpanInline code: output.Append("\\texttt{").Append(EscapeText(code.Text)).Append('}'); break;
                case UnderlineInline underline: output.Append("\\underline{").Append(EscapeText(underline.Text)).Append('}'); break;
                case StrikethroughInline strike:
                    state.Packages.Add("ulem");
                    output.Append("\\sout{").Append(EscapeText(strike.Text)).Append('}');
                    break;
                case StrikethroughSequenceInline strike:
                    state.Packages.Add("ulem");
                    output.Append("\\sout{").Append(Convert(strike.Inlines, state, diagnostics, owner)).Append('}');
                    break;
                case HighlightInline highlight: output.Append("\\textbf{").Append(EscapeText(highlight.Text)).Append('}'); break;
                case SuperscriptInline superscript: output.Append("\\textsuperscript{").Append(EscapeText(superscript.Text)).Append('}'); break;
                case SuperscriptSequenceInline superscript:
                    output.Append("\\textsuperscript{").Append(Convert(superscript.Inlines, state, diagnostics, owner)).Append('}');
                    break;
                case SubscriptInline subscript: output.Append("\\textsubscript{").Append(EscapeText(subscript.Text)).Append('}'); break;
                case SubscriptSequenceInline subscript:
                    output.Append("\\textsubscript{").Append(Convert(subscript.Inlines, state, diagnostics, owner)).Append('}');
                    break;
                case LinkInline link:
                    state.Packages.Add("hyperref");
                    if (link.Url.StartsWith("#", StringComparison.Ordinal)) {
                        output.Append("\\ref{").Append(NormalizeLabel(link.Url.Substring(1), diagnostics, owner)).Append('}');
                    } else {
                        output.Append("\\href{").Append(EscapeArgument(link.Url)).Append("}{")
                            .Append(link.LabelInlines == null ? EscapeText(link.Text) : Convert(link.LabelInlines, state, diagnostics, owner)).Append('}');
                    }
                    break;
                case ImageInline image:
                    state.Packages.Add("graphicx");
                    output.Append("\\includegraphics{").Append(EscapeArgument(image.Src)).Append('}');
                    break;
                case FootnoteRefInline footnote: output.Append("\\footnote{").Append(EscapeText(footnote.Label)).Append('}'); break;
                case HardBreakInline: output.Append("\\\\").Append(state.LineEnding); break;
                case SoftBreakInline: output.Append(state.LineEnding); break;
                case HtmlRawInline html:
                    output.Append("\\texttt{").Append(EscapeText(html.Html)).Append('}');
                    diagnostics.Add(new LatexMarkdownConversionDiagnostic(
                        "MDLATEX103", LatexMarkdownConversionOutcome.Simplified, "raw-html-inline",
                        "Raw HTML was retained as monospaced visible text.", null, owner.SourceSpan));
                    break;
                default:
                    output.Append("\\texttt{").Append(EscapeText(RenderFallback(node))).Append('}');
                    diagnostics.Add(new LatexMarkdownConversionDiagnostic(
                        "MDLATEX109", LatexMarkdownConversionOutcome.SourceFallback, node.GetType().Name,
                        "Unsupported Markdown inline retained as monospaced Markdown source.", null, owner.SourceSpan));
                    break;
            }
        }
        return output.ToString();
    }

    internal static string EscapeText(string value) {
        var output = new StringBuilder(value.Length);
        for (int index = 0; index < value.Length; index++) {
            switch (value[index]) {
                case '\\': output.Append("\\textbackslash{}"); break;
                case '{': output.Append("\\{"); break;
                case '}': output.Append("\\}"); break;
                case '%': output.Append("\\%"); break;
                case '$': output.Append("\\$"); break;
                case '&': output.Append("\\&"); break;
                case '#': output.Append("\\#"); break;
                case '_': output.Append("\\_"); break;
                case '~': output.Append("\\textasciitilde{}"); break;
                case '^': output.Append("\\textasciicircum{}"); break;
                default: output.Append(value[index]); break;
            }
        }
        return output.ToString();
    }

    internal static string EscapeArgument(string value) => EscapeText(value);

    internal static string EscapeLabel(string value) {
        var output = new StringBuilder(value.Length);
        for (int index = 0; index < value.Length; index++) {
            char current = value[index];
            if ((current >= 'a' && current <= 'z') || (current >= 'A' && current <= 'Z') ||
                (current >= '0' && current <= '9') || current == ':' || current == '.' || current == '_' || current == '-') {
                output.Append(current);
            } else {
                output.Append('_').Append(((int)current).ToString("X4", System.Globalization.CultureInfo.InvariantCulture)).Append('_');
            }
        }
        return output.ToString();
    }

    private static string NormalizeLabel(
        string value,
        List<LatexMarkdownConversionDiagnostic> diagnostics,
        MarkdownObject owner) {
        string normalized = EscapeLabel(value);
        if (!string.Equals(value, normalized, StringComparison.Ordinal)) {
            diagnostics.Add(new LatexMarkdownConversionDiagnostic(
                "MDLATEX031", LatexMarkdownConversionOutcome.Simplified, "label",
                "A Markdown identifier was encoded to a TeX-safe label.", null, owner.SourceSpan));
        }
        return normalized;
    }

    private static string RenderFallback(IMarkdownInline node) {
        var sequence = new InlineSequence { AutoSpacing = false };
        sequence.AddRaw(node);
        return MarkdownDoc.Create().Add(new ParagraphBlock(sequence)).ToMarkdown();
    }
}

internal sealed class ConversionState {
    internal ConversionState(string lineEnding) { LineEnding = lineEnding; }
    internal string LineEnding { get; }
    internal HashSet<string> Packages { get; } = new HashSet<string>(StringComparer.Ordinal);
    internal HashSet<string> TheoremEnvironments { get; } = new HashSet<string>(StringComparer.Ordinal);
    internal bool TitleConsumed { get; set; }
}
