namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    internal sealed class ParagraphParser : IMarkdownBlockParser {
        public bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
            if (!options.Paragraphs) return false;
            // Paragraph begins when none of the other block starters match.
            if (IsAtxHeading(lines[i], out _, out _) ||
                IsCodeFenceOpen(lines[i], out _, out _, out _) ||
                StartsTable(lines, i) ||
                IsParagraphInterruptingUnorderedListLine(lines[i]) ||
                IsOrderedListLine(lines[i], out _, out _) ||
                (options.Callouts && IsCalloutHeader(lines[i], out _, out _)) ||
                IsQuoteStarter(lines[i]) ||
                IsReferenceLinkDefinitionStarter(lines, i, options) ||
                IsFootnoteDefinitionStarter(lines[i], options) ||
                (options.StandaloneImageBlocks && IsImageLine(lines[i]))) return false;

            var sb = new StringBuilder();
            int j = i;
            bool prevHard = false;
            while (j < lines.Length && !string.IsNullOrWhiteSpace(lines[j]) &&
                   !IsAtxHeading(lines[j], out _, out _) &&
                   !IsCodeFenceOpen(lines[j], out _, out _, out _) &&
                   !StartsTable(lines, j) &&
                   !IsParagraphInterruptingUnorderedListLine(lines[j]) &&
                   !IsParagraphInterruptingOrderedListLine(lines[j]) &&
                   (!options.Callouts || !IsCalloutHeader(lines[j], out _, out _)) &&
                   !IsQuoteStarter(lines[j]) &&
                   !IsReferenceLinkDefinitionStarter(lines, j, options) &&
                   !IsFootnoteDefinitionStarter(lines[j], options) &&
                   !(options.StandaloneImageBlocks && IsImageLine(lines[j]))) {
                var raw = lines[j];
                bool hard = EndsWithTwoSpaces(raw);
                var trimmed = raw.TrimEnd();
                trimmed = ConsumeTrailingBackslashHardBreak(trimmed, options, out bool slashHard);
                hard = hard || slashHard;
                if (j > i) sb.Append(prevHard ? "\n" : " ");
                sb.Append(trimmed);
                prevHard = hard;
                j++;
            }
            if (sb.Length == 0) return false;
            var paragraphLines = new List<string>(j - i);
            for (var lineIndex = i; lineIndex < j; lineIndex++) {
                paragraphLines.Add(lines[lineIndex]);
            }

            var (text, sourceMap) = JoinParagraphLinesWithSourceMap(paragraphLines, state.SourceLineOffset + i, options, state);
            doc.Add(new ParagraphBlock(ParseInlines(text, options, state, sourceMap)));
            i = j; return true;
        }

        private static bool IsFootnoteDefinitionStarter(string line, MarkdownReaderOptions options) {
            if (options?.Footnotes != true || string.IsNullOrWhiteSpace(line)) {
                return false;
            }

            int leading = 0;
            while (leading < line.Length && line[leading] == ' ') {
                leading++;
            }

            if (leading >= 4 || (leading < line.Length && line[leading] == '\t')) {
                return false;
            }

            var trimmed = line.TrimStart();
            if (!(trimmed.Length > 4 && trimmed[0] == '[' && trimmed[1] == '^')) {
                return false;
            }

            int rb = trimmed.IndexOf(']');
            return rb >= 2
                   && rb + 1 < trimmed.Length
                   && trimmed[rb + 1] == ':';
        }

        private static bool IsReferenceLinkDefinitionStarter(string[] lines, int index, MarkdownReaderOptions options) {
            return TryParseReferenceLinkDefinition(lines, index, options, out _, out _, out _, out _);
        }

        private static bool EndsWithTwoSpaces(string s) {
            if (string.IsNullOrEmpty(s)) return false;
            int n = s.Length - 1;
            int count = 0;
            while (n >= 0 && s[n] == ' ') { count++; n--; if (count >= 2) return true; }
            return false;
        }
    }

    private static bool IsQuoteStarter(string line) {
        if (string.IsNullOrEmpty(line)) return false;
        var t = line.TrimStart();
        return t.StartsWith(">");
    }

    private static string ExpandReferenceLinks(string text, MarkdownReaderState state) {
        if (state == null || state.LinkRefs.Count == 0 || string.IsNullOrEmpty(text)) return text;
        var sb = new System.Text.StringBuilder(text.Length + 16);
        int pos = 0;
        while (pos < text.Length) {
            if (text[pos] == '[') {
                int rb = FindMatchingBracket(text, pos);
                if (rb > pos + 1) {
                    // collapsed: [text][]
                    if (rb + 2 < text.Length && text[rb + 1] == '[' && text[rb + 2] == ']') {
                        var lbl = text.Substring(pos + 1, rb - (pos + 1));
                        var key = NormalizeReferenceLabel(lbl);
                        if (state.LinkRefs.TryGetValue(key, out var defc)) {
                            sb.Append('[').Append(lbl).Append(']')
                              .Append('(').Append(FormatExpandedReferenceDestination(defc.Url));
                            if (!string.IsNullOrEmpty(defc.Title)) sb.Append(' ').Append('"').Append(defc.Title).Append('"');
                            sb.Append(')');
                            pos = rb + 3; continue;
                        }
                    }
                    // full: [text][label]
                    if (rb + 1 < text.Length && text[rb + 1] == '[') {
                        int rb2 = FindMatchingBracket(text, rb + 1);
                        if (rb2 > rb + 2) {
                            var textLbl = text.Substring(pos + 1, rb - (pos + 1));
                            var refLbl = text.Substring(rb + 2, rb2 - (rb + 2));
                            var key = NormalizeReferenceLabel(refLbl);
                            if (state.LinkRefs.TryGetValue(key, out var def)) {
                                sb.Append('[').Append(textLbl).Append(']')
                                  .Append('(').Append(FormatExpandedReferenceDestination(def.Url));
                                if (!string.IsNullOrEmpty(def.Title)) sb.Append(' ').Append('"').Append(def.Title).Append('"');
                                sb.Append(')');
                                pos = rb2 + 1; continue;
                            }
                        }
                    }
                    // shortcut: [label]
                    if (!(rb + 1 < text.Length && (text[rb + 1] == '(' || text[rb + 1] == '['))) {
                        var lbls = text.Substring(pos + 1, rb - (pos + 1));
                        var key = NormalizeReferenceLabel(lbls);
                        if (state.LinkRefs.TryGetValue(key, out var defs)) {
                            sb.Append('[').Append(lbls).Append(']')
                              .Append('(').Append(FormatExpandedReferenceDestination(defs.Url));
                            if (!string.IsNullOrEmpty(defs.Title)) sb.Append(' ').Append('"').Append(defs.Title).Append('"');
                            sb.Append(')');
                            pos = rb + 1; continue;
                        }
                    }
                }
            }
            sb.Append(text[pos]); pos++;
        }
        return sb.ToString();
    }

    private static string FormatExpandedReferenceDestination(string? url) {
        if (string.IsNullOrEmpty(url)) {
            return string.Empty;
        }

        var value = url!;
        return value.IndexOfAny(new[] { ' ', '\t', '\r', '\n' }) >= 0 ? "<" + value + ">" : value;
    }
}
