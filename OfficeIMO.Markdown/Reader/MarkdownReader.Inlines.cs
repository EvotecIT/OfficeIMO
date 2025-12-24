namespace OfficeIMO.Markdown;

/// <summary>
/// Inline parsing helpers for <see cref="MarkdownReader"/>.
/// </summary>
public static partial class MarkdownReader {
    private static InlineSequence ParseInlines(string text, MarkdownReaderOptions options, MarkdownReaderState? state = null) {
        var seq = new InlineSequence();
        if (string.IsNullOrEmpty(text)) return seq;

        int pos = 0;
        while (pos < text.Length) {
            // Hard break signal encoded by paragraph joiner as a bare '\n'
            if (text[pos] == '\n') { seq.HardBreak(); pos++; continue; }
            // HTML-style line breaks in source (commonly used inside table cells): <br>, <br/>, <br />
            if (options.InlineHtml && text[pos] == '<') {
                const string br = "<br>";
                const string brSelf = "<br/>";
                const string brSelfSpaced = "<br />";
                if (text.Length - pos >= br.Length && string.Compare(text, pos, br, 0, br.Length, StringComparison.OrdinalIgnoreCase) == 0) {
                    seq.HardBreak(); pos += br.Length; continue;
                }
                if (text.Length - pos >= brSelf.Length && string.Compare(text, pos, brSelf, 0, brSelf.Length, StringComparison.OrdinalIgnoreCase) == 0) {
                    seq.HardBreak(); pos += brSelf.Length; continue;
                }
                if (text.Length - pos >= brSelfSpaced.Length && string.Compare(text, pos, brSelfSpaced, 0, brSelfSpaced.Length, StringComparison.OrdinalIgnoreCase) == 0) {
                    seq.HardBreak(); pos += brSelfSpaced.Length; continue;
                }
            }
            // Backslash escape: consume next char literally
            if (text[pos] == '\\' && pos + 1 < text.Length) {
                seq.Text(text[pos + 1].ToString());
                pos += 2; continue;
            }

            // Autolink: http(s)://... until whitespace or closing punct
            if (StartsWithHttp(text, pos, out int urlEnd)) {
                var url = text.Substring(pos, urlEnd - pos);
                seq.Link(url, ResolveUrl(url, options) ?? url, null);
                pos = urlEnd; continue;
            }
            if (text[pos] == '`') {
                // Support multi-backtick code spans: count fence length and find a matching run
                int fenceLen = 0; int k = pos; while (k < text.Length && text[k] == '`') { fenceLen++; k++; }
                int j = k; int run = 0; int matchStart = -1;
                while (j < text.Length) {
                    if (text[j] == '`') { run++; if (run == fenceLen) { matchStart = j - fenceLen + 1; break; } j++; continue; }
                    run = 0; j++;
                }
                if (matchStart >= 0) {
                    int contentStart = pos + fenceLen;
                    int contentLen = matchStart - contentStart;
                    if (contentLen < 0) contentLen = 0;
                    var inner = text.Substring(contentStart, contentLen);
                    // Trim one leading/trailing space when surrounded by spaces per CommonMark? Keep simple: preserve as-is
                    seq.Code(inner);
                    pos = matchStart + fenceLen; continue;
                }
            }

            // Footnote ref [^id] should be recognized before generic link parsing
            if (text[pos] == '[' && pos + 2 < text.Length && text[pos + 1] == '^') {
                int rb = text.IndexOf(']', pos + 2);
                if (rb > pos + 2) { var lab = text.Substring(pos + 2, rb - (pos + 2)); seq.FootnoteRef(lab); pos = rb + 1; continue; }
            }

            if (TryParseImageLink(text, pos, out int consumed, out var alt2, out var img2, out var imgTitle2, out var href2)) {
                seq.ImageLink(alt2, ResolveUrl(img2, options) ?? img2, ResolveUrl(href2, options) ?? href2, imgTitle2); pos += consumed; continue;
            }

            if (text[pos] == '!') {
                // Inline image: ![alt](src "title")
                if (TryParseInlineImage(text, pos, out int consumedImg, out var altImg, out var srcImg, out var titleImg)) {
                    seq.Image(altImg, ResolveUrl(srcImg, options) ?? srcImg, titleImg); pos += consumedImg; continue;
                }
            }
            if (text[pos] == '[') {
                if (state != null && TryParseCollapsedRef(text, pos, out int consumedC, out var lbl2)) {
                    if (state.LinkRefs.TryGetValue(lbl2, out var def2)) seq.Link(lbl2, def2.Url, def2.Title); else seq.Text(text.Substring(pos, consumedC));
                    pos += consumedC; continue;
                }
                if (state != null && TryParseRefLink(text, pos, out int consumedR, out var lbl, out var refLabel)) {
                    if (state.LinkRefs.TryGetValue(refLabel, out var def)) seq.Link(lbl, def.Url, def.Title); else seq.Text(text.Substring(pos, consumedR));
                    pos += consumedR; continue;
                }
                if (state != null && TryParseShortcutRef(text, pos, out int consumedS, out var lbl3)) {
                    if (state.LinkRefs.TryGetValue(lbl3, out var def3)) seq.Link(lbl3, def3.Url, def3.Title); else seq.Text(text.Substring(pos, consumedS));
                    pos += consumedS; continue;
                }
                if (TryParseLink(text, pos, out int consumed2, out var label2, out var href3, out var title2)) { seq.Link(label2, ResolveUrl(href3, options) ?? href3, title2); pos += consumed2; continue; }
            }

            // Combined bold+italic ***text*** or ___text___
            if ((text[pos] == '*' && pos + 2 < text.Length && text[pos + 1] == '*' && text[pos + 2] == '*') ||
                (text[pos] == '_' && pos + 2 < text.Length && text[pos + 1] == '_' && text[pos + 2] == '_')) {
                char m = text[pos];
                int end = text.IndexOf(new string(m, 3), pos + 3, System.StringComparison.Ordinal);
                if (end >= 0) {
                    var inner = text.Substring(pos + 3, end - (pos + 3));
                    seq.BoldItalic(inner);
                    pos = end + 3; continue;
                }
            }

            // Bold **text** or __text__
            if ((text[pos] == '*' && pos + 1 < text.Length && text[pos + 1] == '*') ||
                (text[pos] == '_' && pos + 1 < text.Length && text[pos + 1] == '_')) {
                int end = text.IndexOf("**", pos + 2, StringComparison.Ordinal);
                if (text[pos] == '_') end = text.IndexOf("__", pos + 2, StringComparison.Ordinal);
                if (end >= 0) { seq.Bold(text.Substring(pos + 2, end - (pos + 2))); pos = end + 2; continue; }
            }

            if (text[pos] == '~' && pos + 1 < text.Length && text[pos + 1] == '~') {
                int end = text.IndexOf("~~", pos + 2, StringComparison.Ordinal);
                if (end >= 0) { seq.Strike(text.Substring(pos + 2, end - (pos + 2))); pos = end + 2; continue; }
            }

            if (text[pos] == '_' || text[pos] == '*') {
                int end = text.IndexOf('_', pos + 1);
                if (text[pos] == '*') end = text.IndexOf('*', pos + 1);
                if (end > pos + 1) { seq.Italic(text.Substring(pos + 1, end - pos - 1)); pos = end + 1; continue; }
            }

            if (options.InlineHtml && text[pos] == '<') {
                const string uOpen = "<u>"; const string uClose = "</u>";
                if (text.Substring(pos).StartsWith(uOpen, StringComparison.OrdinalIgnoreCase)) {
                    int end = text.IndexOf(uClose, pos + uOpen.Length, StringComparison.OrdinalIgnoreCase);
                    if (end > 0) { var inner = text.Substring(pos + uOpen.Length, end - (pos + uOpen.Length)); seq.Underline(System.Net.WebUtility.HtmlDecode(inner)); pos = end + uClose.Length; continue; }
                }
            }

            // Footnote ref [^id]
            if (text[pos] == '[' && pos + 2 < text.Length && text[pos + 1] == '^') {
                int rb = text.IndexOf(']', pos + 2);
                if (rb > pos + 2) { var lab = text.Substring(pos + 2, rb - (pos + 2)); seq.FootnoteRef(lab); pos = rb + 1; continue; }
            }

            int start = pos; pos++;
            while (pos < text.Length && !IsPotentialInlineStart(text[pos], options.InlineHtml)) pos++;
            seq.Text(text.Substring(start, pos - start));
        }

        return seq;
    }

    private static bool TryParseRefLink(string text, int start, out int consumed, out string label, out string refLabel) {
        consumed = 0; label = refLabel = string.Empty;
        if (start >= text.Length || text[start] != '[') return false;
        int rb = text.IndexOf(']', start + 1); if (rb < 0) return false;
        if (rb + 1 >= text.Length || text[rb + 1] != '[') return false;
        int rb2 = text.IndexOf(']', rb + 2); if (rb2 < 0) return false;
        label = text.Substring(start + 1, rb - (start + 1));
        refLabel = text.Substring(rb + 2, rb2 - (rb + 2));
        consumed = rb2 - start + 1; return true;
    }

    private static bool TryParseCollapsedRef(string text, int start, out int consumed, out string label) {
        consumed = 0; label = string.Empty;
        if (start >= text.Length || text[start] != '[') return false;
        int rb = text.IndexOf(']', start + 1); if (rb < 0) return false;
        if (rb + 2 >= text.Length || text[rb + 1] != '[' || text[rb + 2] != ']') return false;
        label = text.Substring(start + 1, rb - (start + 1));
        consumed = rb + 3 - start;
        return true;
    }

    private static bool TryParseShortcutRef(string text, int start, out int consumed, out string label) {
        consumed = 0; label = string.Empty;
        if (start >= text.Length || text[start] != '[') return false;
        int rb = text.IndexOf(']', start + 1); if (rb < 0) return false;
        if (rb + 1 < text.Length && (text[rb + 1] == '(' || text[rb + 1] == '[')) return false;
        label = text.Substring(start + 1, rb - (start + 1));
        consumed = rb + 1 - start;
        return true;
    }

    private static string? ResolveUrl(string url, MarkdownReaderOptions? options) {
        if (string.IsNullOrWhiteSpace(url)) return null;
        if (url.StartsWith("http://", StringComparison.OrdinalIgnoreCase) || url.StartsWith("https://", StringComparison.OrdinalIgnoreCase)) return url;
        if (url.StartsWith("//")) return url;
        if (url.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase) || url.StartsWith("data:", StringComparison.OrdinalIgnoreCase)) return url;
        if (url.StartsWith("#")) return url;
        var baseUri = options?.BaseUri;
        if (!string.IsNullOrWhiteSpace(baseUri)) {
            try {
                var resolved = new Uri(new Uri(baseUri, UriKind.Absolute), url);
                if (!resolved.Scheme.Equals("http", StringComparison.OrdinalIgnoreCase) &&
                    !resolved.Scheme.Equals("https", StringComparison.OrdinalIgnoreCase)) {
                    return url; // refuse non-http(s) schemes
                }
                return resolved.ToString();
            }
            catch (UriFormatException) { /* invalid base or relative path; keep original */ }
        }
        return url; // leave as-is
    }



    private static readonly bool[] PotentialInlineStartLookup = CreatePotentialInlineStartLookup();

    private static bool[] CreatePotentialInlineStartLookup() {
        var lookup = new bool[128];
        lookup['['] = true;
        lookup['!'] = true;
        lookup['`'] = true;
        lookup['*'] = true;
        lookup['_'] = true;
        lookup['~'] = true;
        return lookup;
    }

    private static bool IsPotentialInlineStart(char c, bool allowInlineHtml) {
        if (c < PotentialInlineStartLookup.Length && PotentialInlineStartLookup[c]) return true;
        return allowInlineHtml && c == '<';
    }

    private static bool TryParseLink(string text, int start, out int consumed, out string label, out string href, out string? title) {
        consumed = 0; label = href = string.Empty; title = null;
        if (start >= text.Length || text[start] != '[') return false;
        int labelEnd = text.IndexOf(']', start + 1);
        if (labelEnd < 0) return false;
        int parenOpen = (labelEnd + 1 < text.Length && text[labelEnd + 1] == '(') ? labelEnd + 1 : -1;
        if (parenOpen < 0) return false;
        int parenClose = FindMatchingParen(text, parenOpen);
        if (parenClose < 0) return false;
        label = text.Substring(start + 1, labelEnd - (start + 1));
        string inner = text.Substring(parenOpen + 1, parenClose - (parenOpen + 1)).Trim();
        href = inner;
        // Optional title: separated by space and in quotes. Find first quote if any.
        int q = inner.IndexOf('"');
        if (q >= 0) {
            href = inner.Substring(0, q).Trim();
            int q2 = inner.LastIndexOf('"');
            if (q2 > q) title = inner.Substring(q + 1, q2 - q - 1);
        } else {
            // no quotes; trim trailing spaces
            href = href.Trim();
        }
        consumed = parenClose - start + 1;
        return true;
    }

    private static bool TryParseImageLink(string text, int start, out int consumed, out string alt, out string img, out string? imgTitle, out string href) {
        consumed = 0; alt = img = href = string.Empty; imgTitle = null;
        if (start >= text.Length || text[start] != '[') return false;
        if (start + 1 >= text.Length || text[start + 1] != '!') return false;
        if (start + 2 >= text.Length || text[start + 2] != '[') return false;
        int altEnd = text.IndexOf(']', start + 3);
        if (altEnd < 0) return false;
        if (altEnd + 1 >= text.Length || text[altEnd + 1] != '(') return false;
        int imgClose = FindMatchingParen(text, altEnd + 1);
        if (imgClose < 0) return false;
        alt = text.Substring(start + 3, altEnd - (start + 3));
        string inner = text.Substring(altEnd + 2, imgClose - (altEnd + 2)).Trim();
        int space = inner.IndexOf(' ');
        if (space < 0) { img = inner; } else {
            img = inner.Substring(0, space).Trim();
            string rest = inner.Substring(space).Trim();
            if (rest.Length >= 2 && rest[0] == '"' && rest[rest.Length - 1] == '"') imgTitle = rest.Substring(1, rest.Length - 2);
        }
        int closeBracket = (imgClose + 1 < text.Length) ? text.IndexOf(']', imgClose + 1) : -1;
        if (closeBracket != imgClose + 1) return false;
        int parenOpen2 = (closeBracket + 1 < text.Length && text[closeBracket + 1] == '(') ? closeBracket + 1 : -1;
        if (parenOpen2 != closeBracket + 1) return false;
        int parenClose2 = FindMatchingParen(text, parenOpen2);
        if (parenClose2 < 0) return false;
        href = text.Substring(parenOpen2 + 1, parenClose2 - (parenOpen2 + 1));
        consumed = parenClose2 - start + 1;
        return true;
    }

    private static bool TryParseInlineImage(string text, int start, out int consumed, out string alt, out string src, out string? title) {
        consumed = 0; alt = src = string.Empty; title = null;
        if (start + 1 >= text.Length || text[start] != '!' || text[start + 1] != '[') return false;
        int altEnd = text.IndexOf(']', start + 2);
        if (altEnd < 0) return false;
        if (altEnd + 1 >= text.Length || text[altEnd + 1] != '(') return false;
        int parenClose = FindMatchingParen(text, altEnd + 1);
        if (parenClose < 0) return false;
        alt = text.Substring(start + 2, altEnd - (start + 2));
        string inner = text.Substring(altEnd + 2, parenClose - (altEnd + 2)).Trim();
        int q = inner.IndexOf('"');
        if (q >= 0) { src = inner.Substring(0, q).Trim(); int q2 = inner.LastIndexOf('"'); if (q2 > q) title = inner.Substring(q + 1, q2 - q - 1); } else { src = inner.Trim(); }
        consumed = parenClose - start + 1;
        return true;
    }

    private static int FindMatchingParen(string text, int openIndex) {
        int depth = 0; bool inQuotes = false;
        for (int i = openIndex; i < text.Length; i++) {
            char c = text[i];
            if (c == '"') { inQuotes = !inQuotes; continue; }
            if (inQuotes) continue;
            if (c == '(') { depth++; continue; }
            if (c == ')') { depth--; if (depth == 0) return i; continue; }
        }
        return -1;
    }

    private static bool StartsWithHttp(string text, int start, out int end) {
        end = start;
        if (start + 7 > text.Length) return false;
        var rem = text.Substring(start);
        if (!(rem.StartsWith("http://") || rem.StartsWith("https://"))) return false;
        int i = start;
        while (i < text.Length) {
            char c = text[i];
            if (char.IsWhiteSpace(c)) break;
            if (c == ')' || c == ']' || c == '<') break;
            i++;
        }
        // Trim trailing punctuation commonly outside URLs
        while (i > start && (text[i - 1] == '.' || text[i - 1] == ',' || text[i - 1] == ';' || text[i - 1] == ':')) i--;
        end = i; return end > start + 7;
    }

    /// <summary>
    /// Parses a single line of Markdown inline content into a typed <see cref="InlineSequence"/>.
    /// This helper is exposed to allow other components (e.g., Word converter) to interpret
    /// inline markup in contexts like table cells where we currently store raw strings.
    /// </summary>
    /// <param name="text">Inline Markdown text.</param>
    /// <param name="options">Reader options controlling inline interpretation.</param>
    /// <returns>Parsed sequence of inline nodes.</returns>
    public static InlineSequence ParseInlineText(string? text, MarkdownReaderOptions? options = null) => ParseInlines(text ?? string.Empty, options ?? new MarkdownReaderOptions(), null);
}
