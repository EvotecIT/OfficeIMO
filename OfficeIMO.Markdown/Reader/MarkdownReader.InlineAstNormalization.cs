using System.Collections.Generic;
using System.Text;

namespace OfficeIMO.Markdown;

/// <summary>
/// AST-level inline normalization helpers for model/chat oriented markdown quirks.
/// </summary>
public static partial class MarkdownReader {
    private static bool NormalizeInlineSequenceInPlace(InlineSequence? sequence, MarkdownInputNormalizationOptions? options) {
        if (sequence == null || options == null) return false;

        bool normalizeEscapedInlineCode = options.NormalizeEscapedInlineCodeSpans;
        bool normalizeTightStrongBoundaries = options.NormalizeTightStrongBoundaries;
        bool normalizeTightColonSpacing = options.NormalizeTightColonSpacing;
        if (!normalizeEscapedInlineCode && !normalizeTightStrongBoundaries && !normalizeTightColonSpacing) return false;

        var items = sequence.Items;
        if (items == null || items.Count == 0) return false;

        var working = new List<object>(items.Count);
        bool changed = false;

        for (int i = 0; i < items.Count; i++) {
            var node = items[i];
            if (node == null) continue;
            if (NormalizeNestedInlineNode(node, options)) changed = true;
            working.Add(node);
        }

        if (normalizeEscapedInlineCode && TryRewriteEscapedInlineCodeSpans(working, out var escapedCodeRewritten)) {
            working = escapedCodeRewritten;
            changed = true;
        }

        if (normalizeTightStrongBoundaries && TryInsertMissingStrongBoundarySpaces(working)) {
            changed = true;
        }

        if (normalizeTightColonSpacing && TryInsertMissingColonBoundarySpaces(working)) {
            changed = true;
        }

        if (changed) {
            sequence.ReplaceItems(CoalesceAdjacentTextRuns(working));
        }

        return changed;
    }

    private static bool NormalizeNestedInlineNode(object node, MarkdownInputNormalizationOptions options) {
        if (node is BoldSequenceInline bold) return NormalizeInlineSequenceInPlace(bold.Inlines, options);
        if (node is ItalicSequenceInline italic) return NormalizeInlineSequenceInPlace(italic.Inlines, options);
        if (node is BoldItalicSequenceInline boldItalic) return NormalizeInlineSequenceInPlace(boldItalic.Inlines, options);
        if (node is StrikethroughSequenceInline strike) return NormalizeInlineSequenceInPlace(strike.Inlines, options);
        if (node is LinkInline link && link.LabelInlines != null) return NormalizeInlineSequenceInPlace(link.LabelInlines, options);
        return false;
    }

    private static bool TryRewriteEscapedInlineCodeSpans(List<object> nodes, out List<object> rewritten) {
        rewritten = new List<object>(nodes.Count);
        bool changed = false;

        int i = 0;
        while (i < nodes.Count) {
            if (nodes[i] is TextRun openTick && openTick.Text == "`") {
                int j = i + 1;
                var body = new StringBuilder();
                bool validBody = true;

                while (j < nodes.Count) {
                    if (nodes[j] is TextRun closeTick && closeTick.Text == "`") break;

                    if (nodes[j] is TextRun textRun) {
                        if (!IsEscapedCodeBodyText(textRun.Text)) {
                            validBody = false;
                            break;
                        }

                        body.Append(textRun.Text);
                        j++;
                        continue;
                    }

                    validBody = false;
                    break;
                }

                if (validBody &&
                    j < nodes.Count &&
                    nodes[j] is TextRun finalTick &&
                    finalTick.Text == "`" &&
                    body.Length > 0) {
                    rewritten.Add(new CodeSpanInline(NormalizeCodeSpanContent(body.ToString())));
                    changed = true;
                    i = j + 1;
                    continue;
                }
            }

            rewritten.Add(nodes[i]);
            i++;
        }

        return changed;
    }

    private static bool TryInsertMissingStrongBoundarySpaces(List<object> nodes) {
        bool changed = false;

        for (int i = 0; i < nodes.Count - 1; i++) {
            if (!IsStrongInlineNode(nodes[i])) continue;
            if (nodes[i + 1] is not TextRun textRun) continue;
            if (!NeedsLeadingSpaceAfterStrong(textRun.Text)) continue;

            nodes[i + 1] = new TextRun(" " + textRun.Text);
            changed = true;
        }

        return changed;
    }

    private static bool TryInsertMissingColonBoundarySpaces(List<object> nodes) {
        bool changed = false;

        for (int i = 0; i < nodes.Count; i++) {
            if (nodes[i] is not TextRun textRun) continue;
            if (string.IsNullOrEmpty(textRun.Text) || textRun.Text.IndexOf(':') < 0) continue;

            var normalized = NormalizeTightColonSpacing(textRun.Text);
            if (normalized == textRun.Text) continue;

            nodes[i] = new TextRun(normalized);
            changed = true;
        }

        return changed;
    }

    private static string NormalizeTightColonSpacing(string text) {
        if (string.IsNullOrEmpty(text) || text.IndexOf(':') < 0) return text;

        StringBuilder? builder = null;
        for (int i = 0; i < text.Length; i++) {
            char current = text[i];
            builder?.Append(current);

            if (current != ':' || !ShouldInsertSpaceAfterColon(text, i)) continue;

            builder ??= new StringBuilder(text.Length + 8);
            if (builder.Length == 0) {
                builder.Append(text, 0, i + 1);
            }
            builder.Append(' ');
        }

        return builder?.ToString() ?? text;
    }

    private static bool ShouldInsertSpaceAfterColon(string text, int colonIndex) {
        if (colonIndex <= 0 || colonIndex >= text.Length - 1) return false;

        char previous = text[colonIndex - 1];
        char next = text[colonIndex + 1];

        if (char.IsWhiteSpace(next)) return false;
        if (!char.IsLetter(previous) || !char.IsLetter(next)) return false;

        return true;
    }

    private static bool IsStrongInlineNode(object node) {
        return node is BoldInline || node is BoldSequenceInline;
    }

    private static bool NeedsLeadingSpaceAfterStrong(string? text) {
        if (string.IsNullOrEmpty(text)) return false;
        return char.IsLetterOrDigit(text![0]);
    }

    private static bool IsEscapedCodeBodyText(string? text) {
        if (text == null) return false;
        if (text.IndexOf('`') >= 0) return false;
        if (text.IndexOf('\r') >= 0) return false;
        if (text.IndexOf('\n') >= 0) return false;
        return true;
    }

    private static List<object> CoalesceAdjacentTextRuns(List<object> nodes) {
        if (nodes.Count <= 1) return nodes;

        var compact = new List<object>(nodes.Count);
        StringBuilder? textBuffer = null;

        void FlushTextBuffer() {
            if (textBuffer == null) return;
            compact.Add(new TextRun(textBuffer.ToString()));
            textBuffer = null;
        }

        for (int i = 0; i < nodes.Count; i++) {
            var node = nodes[i];
            if (node is TextRun textRun) {
                textBuffer ??= new StringBuilder();
                textBuffer.Append(textRun.Text);
                continue;
            }

            FlushTextBuffer();
            compact.Add(node);
        }

        FlushTextBuffer();
        return compact;
    }
}
