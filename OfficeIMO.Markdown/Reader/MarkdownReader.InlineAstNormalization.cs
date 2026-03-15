using System.Collections.Generic;
using System.Text;

namespace OfficeIMO.Markdown;

/// <summary>
/// AST-level inline normalization helpers for model/chat oriented markdown quirks.
/// </summary>
public static partial class MarkdownReader {
    internal static bool NormalizeInlineSequenceInPlace(InlineSequence? sequence, MarkdownInputNormalizationOptions? options) {
        if (sequence == null || options == null) return false;

        bool normalizeEscapedInlineCode = options.NormalizeEscapedInlineCodeSpans;
        bool normalizeTightStrongBoundaries = options.NormalizeTightStrongBoundaries;
        bool normalizeTightColonSpacing = options.NormalizeTightColonSpacing;
        bool normalizeTightParentheticalSpacing = options.NormalizeTightParentheticalSpacing;
        if (!normalizeEscapedInlineCode
            && !normalizeTightStrongBoundaries
            && !normalizeTightColonSpacing
            && !normalizeTightParentheticalSpacing) return false;

        var items = sequence.Nodes;
        if (items == null || items.Count == 0) return false;

        var working = new List<IMarkdownInline>(items.Count);
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

        if (normalizeTightParentheticalSpacing && TryInsertMissingParentheticalBoundarySpaces(working)) {
            changed = true;
        }

        if (changed) {
            sequence.ReplaceItems(CoalesceAdjacentTextRuns(working));
        }

        return changed;
    }

    private static bool NormalizeNestedInlineNode(IMarkdownInline node, MarkdownInputNormalizationOptions options) {
        return node is IInlineContainerMarkdownInline container &&
               NormalizeInlineSequenceInPlace(container.NestedInlines, options);
    }

    private static bool TryRewriteEscapedInlineCodeSpans(List<IMarkdownInline> nodes, out List<IMarkdownInline> rewritten) {
        rewritten = new List<IMarkdownInline>(nodes.Count);
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

    private static bool TryInsertMissingStrongBoundarySpaces(List<IMarkdownInline> nodes) {
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

    private static bool TryInsertMissingColonBoundarySpaces(List<IMarkdownInline> nodes) {
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

    private static bool TryInsertMissingParentheticalBoundarySpaces(List<IMarkdownInline> nodes) {
        bool changed = false;

        for (int i = 0; i < nodes.Count; i++) {
            if (nodes[i] is TextRun textRun) {
                var normalized = NormalizeTightParentheticalSpacing(textRun.Text);
                if (normalized != textRun.Text) {
                    nodes[i] = new TextRun(normalized);
                    textRun = (TextRun)nodes[i];
                    changed = true;
                }
            }

            if (i == 0) {
                continue;
            }

            if (!IsStrongInlineNode(nodes[i - 1])) {
                continue;
            }

            if (nodes[i] is not TextRun nextTextRun || !StartsWithTightParenthetical(nextTextRun.Text)) {
                continue;
            }

            nodes[i] = new TextRun(" " + nextTextRun.Text);
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

    private static string NormalizeTightParentheticalSpacing(string text) {
        if (string.IsNullOrEmpty(text) || text.IndexOf('(') < 0) return text;

        StringBuilder? builder = null;
        for (int i = 0; i < text.Length; i++) {
            char current = text[i];
            if (current == '(' && ShouldInsertSpaceBeforeParenthesis(text, i)) {
                builder ??= new StringBuilder(text.Length + 8);
                if (builder.Length == 0) {
                    builder.Append(text, 0, i);
                }
                builder.Append(' ').Append('(');
                continue;
            }

            builder?.Append(current);
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

    private static bool ShouldInsertSpaceBeforeParenthesis(string text, int openParenIndex) {
        if (openParenIndex <= 0 || openParenIndex >= text.Length - 1) return false;

        char previous = text[openParenIndex - 1];
        char next = text[openParenIndex + 1];

        if (!char.IsLetterOrDigit(previous) && previous != ')') return false;
        if (!char.IsLetter(next)) return false;
        if (char.IsWhiteSpace(previous)) return false;

        return text.IndexOf(')', openParenIndex + 1) > openParenIndex + 1;
    }

    private static bool StartsWithTightParenthetical(string? text) {
        if (string.IsNullOrEmpty(text) || text![0] != '(' || text.Length < 3) {
            return false;
        }

        if (!char.IsLetter(text[1])) {
            return false;
        }

        return text.IndexOf(')', 2) > 1;
    }

    private static bool IsStrongInlineNode(IMarkdownInline node) {
        return node is IStrongMarkdownInline;
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

    private static List<IMarkdownInline> CoalesceAdjacentTextRuns(List<IMarkdownInline> nodes) {
        if (nodes.Count <= 1) return nodes;

        var compact = new List<IMarkdownInline>(nodes.Count);
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
