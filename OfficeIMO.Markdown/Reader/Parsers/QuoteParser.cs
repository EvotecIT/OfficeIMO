namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    internal sealed class QuoteParser : IMarkdownBlockParser {
        public bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
            var t = lines[i];
            // Exclude callouts (handled earlier): they start with "> [!"
            var trimmed = t.TrimStart();
            if (!trimmed.StartsWith(">")) return false;
            if (trimmed.StartsWith(">") && trimmed.Length > 1 && trimmed[1] == ' ' && trimmed.Length > 3 && trimmed[2] == '[' && trimmed[3] == '!') return false;

            // Collect contiguous quote lines and un-prefix one ">" level
            var inner = new System.Collections.Generic.List<string>();
            int j = i;
            bool sawQuotedLine = false;
            while (j < lines.Length) {
                var ln = lines[j];
                var ltrim = ln.TrimStart();
                if (ltrim.StartsWith(">")) {
                    // Strip one level
                    if (ltrim.Length >= 2 && ltrim[1] == ' ') inner.Add(ltrim.Substring(2)); else inner.Add(ltrim.Substring(1));
                    sawQuotedLine = true;
                    j++;
                    continue;
                }

                // Lazy continuation: allow a non-quoted line to continue a blockquote paragraph
                // until a blank line followed by a non-quoted line ends the blockquote.
                if (sawQuotedLine) {
                    if (string.IsNullOrWhiteSpace(ln)) {
                        int peek = j + 1;
                        if (peek >= lines.Length) break;
                        var nextTrim = (lines[peek] ?? string.Empty).TrimStart();
                        if (!nextTrim.StartsWith(">")) break;
                        inner.Add(string.Empty);
                        j++;
                        continue;
                    }

                    // Only continue lazily when both sides look like paragraph content.
                    // A non-quoted list/item/code starter should end the blockquote instead of being swallowed into it.
                    if (inner.Count > 0) {
                        if (LooksLikeParagraphLine(inner, inner.Count - 1, options)) {
                            if (!TryNormalizeQuoteLazyContinuationLine(lines, j, options, out var normalizedLazyLine)) break;

                            inner.Add(normalizedLazyLine);
                            j++;
                            continue;
                        }

                        if (TryNormalizeQuoteLazyContinuationAfterListItem(inner[inner.Count - 1], lines, j, options, out var normalizedListLazyLine)) {
                            inner.Add(normalizedListLazyLine);
                            j++;
                            continue;
                        }
                    }

                    break;
                }

                break;
            }
            // Recursively parse inner content as a separate document
            var nestedOptions = CloneOptionsWithoutFrontMatter(options);
            var nestedState = CloneState(state);
            var syntaxChildren = new System.Collections.Generic.List<MarkdownSyntaxNode>();
            var innerDoc = ParseInternal(string.Join("\n", inner), nestedOptions, nestedState, allowFrontMatter: false, syntaxChildren, lineOffset: state.SourceLineOffset + i);
            var qb = new QuoteBlock();
            foreach (var b in innerDoc.Blocks) qb.Children.Add(b);
            qb.SyntaxChildren = syntaxChildren;
            doc.Add(qb); i = j; return true;
        }
    }

    private static bool LooksLikeParagraphLine(IReadOnlyList<string>? lines, int index, MarkdownReaderOptions options) {
        if (lines == null || index < 0 || index >= lines.Count) return false;
        var line = lines[index] ?? string.Empty;
        if (string.IsNullOrWhiteSpace(line)) return false;
        if (CountLeadingSpaces(line) >= 4) return false;

        var t = line.TrimStart();

        // Block starters we do not want to lazily continue after.
        if (t.StartsWith(">")) return false;
        if (IsAtxHeading(t, out _, out _)) return false;
        if (LooksLikeHr(t)) return false;
        if (IsCodeFenceOpen(t, out _, out _, out _)) return false;
        if (LooksLikeTableRow(t)) return false;
        if (IsUnorderedListLine(t, out _, out _, out _)) return false;
        if (IsParagraphInterruptingOrderedListLine(t)) return false;
        if (ShouldTreatAsDefinitionLine(lines, index, options)) return false;
        if (IsCalloutHeader("> " + t, out _, out _)) return false; // callout marker is quote-prefixed in source

        return true;
    }

    private static bool TryNormalizeQuoteLazyContinuationLine(IReadOnlyList<string>? lines, int index, MarkdownReaderOptions options, out string normalized) {
        var source = lines != null && index >= 0 && index < lines.Count ? (lines[index] ?? string.Empty) : string.Empty;
        normalized = source;
        if (string.IsNullOrWhiteSpace(source)) return false;

        int leadingSpaces = CountLeadingSpaces(source);
        if (leadingSpaces == 0) {
            return LooksLikeParagraphLine(lines, index, options);
        }

        if (leadingSpaces > 4) {
            return false;
        }

        var trimmed = source.TrimStart();
        if (trimmed.Length == 0) return false;
        if (trimmed.StartsWith(">")) return false;
        if (IsAtxHeading(trimmed, out _, out _)) return false;
        if (LooksLikeHr(trimmed)) return false;
        if (IsCodeFenceOpen(trimmed, out _, out _, out _)) return false;
        if (LooksLikeTableRow(trimmed)) return false;
        if (ShouldTreatAsDefinitionLine(lines, index, options)) return false;
        if (IsCalloutHeader("> " + trimmed, out _, out _)) return false;

        if (IsUnorderedListLine(trimmed, out _, out _, out _) || IsParagraphInterruptingOrderedListLine(trimmed)) {
            normalized = "\\" + trimmed;
            return true;
        }

        normalized = trimmed;
        return true;
    }

    private static bool TryNormalizeQuoteLazyContinuationAfterListItem(string? previousLine, IReadOnlyList<string>? lines, int index, MarkdownReaderOptions options, out string normalized) {
        normalized = string.Empty;
        if (string.IsNullOrWhiteSpace(previousLine)) return false;
        if (!TryNormalizeQuoteLazyContinuationLine(lines, index, options, out var normalizedLazyLine)) return false;

        var previous = previousLine!;
        if (!IsUnorderedListLine(previous, out _, out _, out _) &&
            !IsOrderedListLine(previous, out _, out _, out _)) {
            return false;
        }

        int continuationIndent = GetListContinuationIndent(previous);
        normalized = new string(' ', Math.Max(continuationIndent, 1)) + normalizedLazyLine;
        return true;
    }
}
