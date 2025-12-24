namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    internal sealed class ParagraphParser : IMarkdownBlockParser {
        public bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
            if (!options.Paragraphs) return false;
            // Paragraph begins when none of the other block starters match.
            if (IsAtxHeading(lines[i], out _, out _) ||
                IsCodeFenceOpen(lines[i], out _, out _) ||
                StartsTable(lines, i) ||
                IsUnorderedListLine(lines[i], out _, out _, out _) ||
                IsOrderedListLine(lines[i], out _, out _) ||
                IsCalloutHeader(lines[i], out _, out _) ||
                IsQuoteStarter(lines[i]) ||
                IsImageLine(lines[i])) return false;

            var sb = new StringBuilder();
            int j = i;
            while (j < lines.Length && !string.IsNullOrWhiteSpace(lines[j]) &&
                   !IsAtxHeading(lines[j], out _, out _) &&
                   !IsCodeFenceOpen(lines[j], out _, out _) &&
                   !StartsTable(lines, j) &&
                   !IsUnorderedListLine(lines[j], out _, out _, out _) &&
                   !IsOrderedListLine(lines[j], out _, out _) &&
                   !IsCalloutHeader(lines[j], out _, out _) &&
                   !IsQuoteStarter(lines[j]) &&
                   !IsImageLine(lines[j])) {
                var raw = lines[j];
                bool hard = EndsWithTwoSpaces(raw);
                var trimmed = raw.TrimEnd();
                if (sb.Length > 0) sb.Append(hard ? "\n" : " ");
                sb.Append(trimmed);
                j++;
            }
            if (sb.Length == 0) return false;
            var text = ExpandReferenceLinks(sb.ToString(), state);
            doc.Add(new ParagraphBlock(ParseInlines(text, options, state)));
            i = j; return true;
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
                int rb = text.IndexOf(']', pos + 1);
                if (rb > pos + 1) {
                    // collapsed: [text][]
                    if (rb + 2 < text.Length && text[rb + 1] == '[' && text[rb + 2] == ']') {
                        var lbl = text.Substring(pos + 1, rb - (pos + 1));
                        if (state.LinkRefs.TryGetValue(lbl, out var defc)) {
                            sb.Append('[').Append(lbl).Append(']')
                              .Append('(').Append(defc.Url);
                            if (!string.IsNullOrEmpty(defc.Title)) sb.Append(' ').Append('"').Append(defc.Title).Append('"');
                            sb.Append(')');
                            pos = rb + 3; continue;
                        }
                    }
                    // full: [text][label]
                    if (rb + 1 < text.Length && text[rb + 1] == '[') {
                        int rb2 = text.IndexOf(']', rb + 2);
                        if (rb2 > rb + 2) {
                            var textLbl = text.Substring(pos + 1, rb - (pos + 1));
                            var refLbl = text.Substring(rb + 2, rb2 - (rb + 2));
                            if (state.LinkRefs.TryGetValue(refLbl, out var def)) {
                                sb.Append('[').Append(textLbl).Append(']')
                                  .Append('(').Append(def.Url);
                                if (!string.IsNullOrEmpty(def.Title)) sb.Append(' ').Append('"').Append(def.Title).Append('"');
                                sb.Append(')');
                                pos = rb2 + 1; continue;
                            }
                        }
                    }
                    // shortcut: [label]
                    if (!(rb + 1 < text.Length && (text[rb + 1] == '(' || text[rb + 1] == '['))) {
                        var lbls = text.Substring(pos + 1, rb - (pos + 1));
                        if (state.LinkRefs.TryGetValue(lbls, out var defs)) {
                            sb.Append('[').Append(lbls).Append(']')
                              .Append('(').Append(defs.Url);
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
}
