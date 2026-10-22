namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
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
