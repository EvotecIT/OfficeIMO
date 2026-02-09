namespace OfficeIMO.Markdown;

/// <summary>
/// Inline parsing helpers for <see cref="MarkdownReader"/>.
/// </summary>
public static partial class MarkdownReader {
    private static InlineSequence ParseInlines(string text, MarkdownReaderOptions options, MarkdownReaderState? state = null) {
        var seq = new InlineSequence { AutoSpacing = false };
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
            // Backslash escape (CommonMark-ish): only escape punctuation we care about so that Windows paths like
            // "C:\Support\GitHub" keep their backslashes.
            if (text[pos] == '\\') {
                if (pos + 1 < text.Length) {
                    char next = text[pos + 1];
                    if (IsBackslashEscapable(next)) {
                        seq.Text(next.ToString());
                        pos += 2;
                        continue;
                    }
                }
                seq.Text("\\");
                pos++;
                continue;
            }

            // Autolink: http(s)://... until whitespace or closing punct
            if (options.AutolinkUrls && StartsWithHttp(text, pos, out int urlEnd)) {
                var url = text.Substring(pos, urlEnd - pos);
                var resolved = ResolveUrl(url, options);
                if (string.IsNullOrEmpty(resolved)) seq.Text(url);
                else seq.Link(url, resolved!, null);
                pos = urlEnd; continue;
            }

            // Autolink: www.example.com
            if (options.AutolinkWwwUrls && StartsWithWww(text, pos, out int wwwEnd)) {
                var label = text.Substring(pos, wwwEnd - pos);
                var scheme = string.IsNullOrWhiteSpace(options.AutolinkWwwScheme) ? "https://" : options.AutolinkWwwScheme.Trim();
                if (!scheme.EndsWith("://", StringComparison.Ordinal)) scheme = scheme.TrimEnd('/') + "://";
                var href = scheme + label;
                var resolved = ResolveUrl(href, options);
                if (string.IsNullOrEmpty(resolved)) seq.Text(label);
                else seq.Link(label, resolved!, null);
                pos = wwwEnd; continue;
            }

            // Autolink: plain email
            if (options.AutolinkEmails && TryConsumePlainEmail(text, pos, out int emailEnd, out string email)) {
                var href = "mailto:" + email;
                var resolved = ResolveUrl(href, options);
                if (string.IsNullOrEmpty(resolved)) seq.Text(email);
                else seq.Link(email, resolved!, null);
                pos = emailEnd; continue;
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
                    inner = NormalizeCodeSpanContent(inner);
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
                var imgResolved = ResolveUrl(img2, options);
                var hrefResolved = ResolveUrl(href2, options);
                if (string.IsNullOrEmpty(imgResolved) || string.IsNullOrEmpty(hrefResolved)) {
                    // Unsafe URLs: keep content as plain text instead of a clickable linked image.
                    seq.Text(string.IsNullOrEmpty(alt2) ? "image" : alt2);
                } else {
                    seq.ImageLink(alt2, imgResolved!, hrefResolved!, imgTitle2);
                }
                pos += consumed; continue;
            }

            if (text[pos] == '!') {
                // Reference-style image: ![alt][label], ![alt][], or shortcut ![label]
                if (state != null && TryParseReferenceImage(text, pos, out int consumedRefImg, out var altRef, out var refLabel)) {
                    var key = NormalizeReferenceLabel(refLabel);
                    if (state.LinkRefs.TryGetValue(key, out var defImg)) {
                        var resolved = ResolveUrl(defImg.Url, options);
                        if (string.IsNullOrEmpty(resolved)) {
                            seq.Text(string.IsNullOrEmpty(altRef) ? "image" : altRef);
                        } else {
                            seq.Image(altRef, resolved!, defImg.Title);
                        }
                    } else {
                        // Preserve literal syntax when the definition is missing.
                        seq.Text(text.Substring(pos, consumedRefImg));
                    }
                    pos += consumedRefImg; continue;
                }

                // Inline image: ![alt](src "title")
                if (TryParseInlineImage(text, pos, out int consumedImg, out var altImg, out var srcImg, out var titleImg)) {
                    var srcResolved = ResolveUrl(srcImg, options);
                    if (string.IsNullOrEmpty(srcResolved)) {
                        seq.Text(string.IsNullOrEmpty(altImg) ? "image" : altImg);
                    } else {
                        seq.Image(altImg, srcResolved!, titleImg);
                    }
                    pos += consumedImg; continue;
                }
            }

            // Angle-bracket autolinks: <https://example.com> and <user@example.com>
            if (text[pos] == '<' && TryParseAngleAutolink(text, pos, out int consumedAngle, out var labelAngle, out var hrefAngle)) {
                var resolved = ResolveUrl(hrefAngle, options);
                if (string.IsNullOrEmpty(resolved)) {
                    seq.Text(text.Substring(pos, consumedAngle));
                } else {
                    seq.Link(labelAngle, resolved!, null);
                }
                pos += consumedAngle;
                continue;
            }
            if (text[pos] == '[') {
                if (state != null && TryParseCollapsedRef(text, pos, out int consumedC, out var lbl2)) {
                    var key = NormalizeReferenceLabel(lbl2);
                    if (state.LinkRefs.TryGetValue(key, out var def2)) {
                        var resolved = ResolveUrl(def2.Url, options);
                        if (string.IsNullOrEmpty(resolved)) seq.Text(lbl2);
                        else seq.Link(lbl2, resolved!, def2.Title);
                    } else seq.Text(text.Substring(pos, consumedC));
                    pos += consumedC; continue;
                }
                if (state != null && TryParseRefLink(text, pos, out int consumedR, out var lbl, out var refLabel)) {
                    var key = NormalizeReferenceLabel(refLabel);
                    if (state.LinkRefs.TryGetValue(key, out var def)) {
                        var resolved = ResolveUrl(def.Url, options);
                        if (string.IsNullOrEmpty(resolved)) seq.Text(lbl);
                        else seq.Link(lbl, resolved!, def.Title);
                    } else seq.Text(text.Substring(pos, consumedR));
                    pos += consumedR; continue;
                }
                if (state != null && TryParseShortcutRef(text, pos, out int consumedS, out var lbl3)) {
                    var key = NormalizeReferenceLabel(lbl3);
                    if (state.LinkRefs.TryGetValue(key, out var def3)) {
                        var resolved = ResolveUrl(def3.Url, options);
                        if (string.IsNullOrEmpty(resolved)) seq.Text(lbl3);
                        else seq.Link(lbl3, resolved!, def3.Title);
                    } else seq.Text(text.Substring(pos, consumedS));
                    pos += consumedS; continue;
                }
                if (TryParseLink(text, pos, out int consumed2, out var label2, out var href3, out var title2)) {
                    // Allow empty href: commonly used as placeholder or to be filled by the host.
                    if (string.IsNullOrWhiteSpace(href3)) {
                        seq.Link(label2, string.Empty, title2);
                    } else {
                        var hrefResolved = ResolveUrl(href3, options);
                        if (string.IsNullOrEmpty(hrefResolved)) {
                            // Unsafe URLs: keep the label as plain text instead of producing an <a href="...">.
                            seq.Text(label2);
                        } else {
                            seq.Link(label2, hrefResolved!, title2);
                        }
                    }
                    pos += consumed2; continue;
                }
            }

            // Combined bold+italic ***text*** or ___text___
            if ((text[pos] == '*' && pos + 2 < text.Length && text[pos + 1] == '*' && text[pos + 2] == '*') ||
                (text[pos] == '_' && pos + 2 < text.Length && text[pos + 1] == '_' && text[pos + 2] == '_')) {
                char m = text[pos];
                if (m == '_' && IsIntrawordDelimiter(text, pos, 3)) {
                    // Do not treat intraword underscores as emphasis markers (e.g., foo___bar___baz).
                } else {
                    int end = text.IndexOf(new string(m, 3), pos + 3, System.StringComparison.Ordinal);
                    if (end >= 0) {
                        var inner = text.Substring(pos + 3, end - (pos + 3));
                        seq.AddRaw(new BoldItalicSequenceInline(ParseInlines(inner, options, state)));
                        pos = end + 3; continue;
                    }
                }
            }

            // Bold **text** or __text__
            if ((text[pos] == '*' && pos + 1 < text.Length && text[pos + 1] == '*') ||
                (text[pos] == '_' && pos + 1 < text.Length && text[pos + 1] == '_')) {
                if (text[pos] == '_' && IsIntrawordDelimiter(text, pos, 2)) {
                    // do not parse intraword "__" as emphasis
                } else {
                    int end = text.IndexOf("**", pos + 2, StringComparison.Ordinal);
                    if (text[pos] == '_') end = text.IndexOf("__", pos + 2, StringComparison.Ordinal);
                    if (end >= 0) {
                        var inner = text.Substring(pos + 2, end - (pos + 2));
                        seq.AddRaw(new BoldSequenceInline(ParseInlines(inner, options, state)));
                        pos = end + 2; continue;
                    }
                }
            }

            if (text[pos] == '~' && pos + 1 < text.Length && text[pos + 1] == '~') {
                int end = text.IndexOf("~~", pos + 2, StringComparison.Ordinal);
                if (end >= 0) {
                    var inner = text.Substring(pos + 2, end - (pos + 2));
                    seq.AddRaw(new StrikethroughSequenceInline(ParseInlines(inner, options, state)));
                    pos = end + 2; continue;
                }
            }

            if (text[pos] == '_' || text[pos] == '*') {
                if (text[pos] == '_' && IsIntrawordDelimiter(text, pos, 1)) {
                    // do not parse intraword "_" as emphasis
                } else {
                    char m = text[pos];
                    int end = FindSingleEmphasisEnd(text, pos, m);
                    if (end > pos + 1) {
                        var inner = text.Substring(pos + 1, end - pos - 1);
                        seq.AddRaw(new ItalicSequenceInline(ParseInlines(inner, options, state)));
                        pos = end + 1; continue;
                    }
                }
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
            while (pos < text.Length && !IsPotentialInlineStart(text[pos], options.InlineHtml)) {
                // Ensure our explicit inline handlers see these characters.
                if (text[pos] == '\n') break;
                if (text[pos] == '\\' && pos + 1 < text.Length && IsBackslashEscapable(text[pos + 1])) break;
                if (text[pos] == '<' && IsAngleAutolinkStart(text, pos)) break;
                if (options.AutolinkUrls && (text[pos] == 'h' || text[pos] == 'H') && StartsWithHttp(text, pos, out _)) break;
                if (options.AutolinkWwwUrls && (text[pos] == 'w' || text[pos] == 'W') && StartsWithWww(text, pos, out _)) break;
                if (options.AutolinkEmails && IsEmailStartChar(text[pos]) && TryConsumePlainEmail(text, pos, out _, out _)) break;
                pos++;
            }
            seq.Text(text.Substring(start, pos - start));
        }

        return seq;
    }

    private static int FindSingleEmphasisEnd(string text, int start, char marker) {
        if (string.IsNullOrEmpty(text)) return -1;
        int i = start + 1;
        while (i < text.Length) {
            i = text.IndexOf(marker, i);
            if (i < 0) return -1;

            // Avoid closing on a delimiter that's part of a longer run (e.g., the '*' in "**bold**").
            if ((i > start + 1 && text[i - 1] == marker) || (i + 1 < text.Length && text[i + 1] == marker)) {
                i++;
                continue;
            }

            // Avoid closing on intraword '_' (e.g., "foo_bar_baz").
            if (marker == '_' && IsIntrawordDelimiter(text, i, 1)) {
                i++;
                continue;
            }

            return i;
        }
        return -1;
    }

    private static string NormalizeCodeSpanContent(string inner) {
        if (inner == null) return string.Empty;

        // Normalize newlines to spaces (CommonMark-like).
        if (inner.IndexOf('\r') >= 0) inner = inner.Replace("\r\n", "\n").Replace("\r", "\n");
        if (inner.IndexOf('\n') >= 0) inner = inner.Replace("\n", " ");

        // Trim a single leading+trailing space if both exist and the content is not all spaces.
        if (inner.Length >= 2 && inner[0] == ' ' && inner[inner.Length - 1] == ' ') {
            bool anyNonSpace = false;
            for (int i = 0; i < inner.Length; i++) {
                if (inner[i] != ' ') { anyNonSpace = true; break; }
            }
            if (anyNonSpace) inner = inner.Substring(1, inner.Length - 2);
        }

        return inner;
    }

    private static bool IsAngleAutolinkStart(string text, int start) {
        if (start < 0 || start >= text.Length) return false;
        if (text[start] != '<') return false;
        return TryParseAngleAutolink(text, start, out _, out _, out _);
    }

    private static bool TryParseAngleAutolink(string text, int start, out int consumed, out string label, out string href) {
        consumed = 0;
        label = href = string.Empty;
        if (start < 0 || start >= text.Length || text[start] != '<') return false;
        int gt = text.IndexOf('>', start + 1);
        if (gt < 0) return false;
        if (gt == start + 1) return false;

        // Disallow whitespace/control inside.
        for (int i = start + 1; i < gt; i++) {
            char c = text[i];
            if (char.IsWhiteSpace(c) || char.IsControl(c)) return false;
        }

        var inner = text.Substring(start + 1, gt - (start + 1));

        // URL form
        if (inner.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
            inner.StartsWith("https://", StringComparison.OrdinalIgnoreCase)) {
            label = inner;
            href = inner;
            consumed = gt - start + 1;
            return true;
        }

        // Email form
        if (LooksLikeEmail(inner)) {
            label = inner;
            href = "mailto:" + inner;
            consumed = gt - start + 1;
            return true;
        }

        return false;
    }

    private static bool LooksLikeEmail(string s) {
        if (string.IsNullOrEmpty(s)) return false;
        int at = s.IndexOf('@');
        if (at <= 0 || at == s.Length - 1) return false;
        // Reject "mailto:" which is a URL form and will be handled above if ever enabled.
        if (s.IndexOf(':') >= 0) return false;

        string domain = s.Substring(at + 1);
        // Require at least one '.' in domain and not at the ends.
        int dot = domain.IndexOf('.');
        if (dot <= 0 || dot == domain.Length - 1) return false;

        // Basic character checks (no spaces/control already enforced by caller).
        for (int i = 0; i < s.Length; i++) {
            char c = s[i];
            if (c == '@') continue;
            if (c == '.' || c == '-' || c == '_' || c == '+') continue;
            if (char.IsLetterOrDigit(c)) continue;
            return false;
        }
        return true;
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
        url = url.Trim();

        // Block scriptable schemes by default.
        if (TryGetScheme(url, out var scheme)) {
            if (options?.DisallowScriptUrls != false) {
                if (scheme.Equals("javascript", StringComparison.OrdinalIgnoreCase) ||
                    scheme.Equals("vbscript", StringComparison.OrdinalIgnoreCase)) {
                    return null;
                }
            }
            if (options?.DisallowFileUrls == true) {
                if (scheme.Equals("file", StringComparison.OrdinalIgnoreCase) || IsWindowsDriveLike(url)) return null;
            }
            if (scheme.Equals("mailto", StringComparison.OrdinalIgnoreCase)) return (options?.AllowMailtoUrls ?? true) ? url : null;
            if (scheme.Equals("data", StringComparison.OrdinalIgnoreCase)) return (options?.AllowDataUrls ?? true) ? url : null;
            // http/https and unknown schemes: keep as-is (host may further restrict)
            return url;
        }

        if (url.StartsWith("//")) return (options?.AllowProtocolRelativeUrls ?? true) ? url : null;
        if (url.StartsWith("#")) return url;
        if (options?.DisallowFileUrls == true && IsWindowsDriveLike(url)) return null;

        var baseUri = options?.BaseUri;
        if (!string.IsNullOrWhiteSpace(baseUri)) {
            try {
                // Legacy behavior: only apply BaseUri when it is http(s), and only resolve into http(s).
                var baseAbs = new Uri(baseUri, UriKind.Absolute);
                if (!baseAbs.Scheme.Equals("http", StringComparison.OrdinalIgnoreCase) &&
                    !baseAbs.Scheme.Equals("https", StringComparison.OrdinalIgnoreCase)) {
                    return url;
                }
                var resolved = new Uri(baseAbs, url);
                if (!resolved.Scheme.Equals("http", StringComparison.OrdinalIgnoreCase) &&
                    !resolved.Scheme.Equals("https", StringComparison.OrdinalIgnoreCase)) {
                    return url;
                }
                return resolved.ToString();
            }
            catch (UriFormatException) { /* invalid base or relative path; keep original */ }
        }

        return url; // relative or unknown: leave as-is
    }

    private static bool TryGetScheme(string url, out string scheme) {
        scheme = string.Empty;
        int colon = url.IndexOf(':');
        if (colon <= 0) return false;
        // If there's a path/query/fragment delimiter before ':', it's not a scheme.
        int slash = url.IndexOfAny(new[] { '/', '?', '#' });
        if (slash >= 0 && slash < colon) return false;
        // URI scheme must start with a letter and be [A-Za-z0-9+.-]*
        char first = url[0];
        if (!char.IsLetter(first)) return false;
        for (int i = 1; i < colon; i++) {
            char c = url[i];
            bool ok = char.IsLetterOrDigit(c) || c == '+' || c == '-' || c == '.';
            if (!ok) return false;
        }
        scheme = url.Substring(0, colon);
        return true;
    }

    private static bool IsWindowsDriveLike(string url) {
        // Treat "C:\..." and "C:/..." as file-like.
        return url.Length >= 3
               && char.IsLetter(url[0])
               && url[1] == ':'
               && (url[2] == '\\' || url[2] == '/');
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

    private static bool IsBackslashEscapable(char c) {
        // CommonMark backslash-escapable punctuation (plus '|' which we want for tables).
        // See: https://spec.commonmark.org/ (backslash escapes). We keep the set small and pragmatic.
        return c switch {
            '\\' => true,
            '`' => true,
            '*' => true,
            '_' => true,
            '{' => true,
            '}' => true,
            '[' => true,
            ']' => true,
            '(' => true,
            ')' => true,
            '#' => true,
            '+' => true,
            '-' => true,
            '.' => true,
            '!' => true,
            '|' => true,
            '>' => true,
            _ => false
        };
    }

    private static bool IsIntrawordDelimiter(string text, int start, int markerLength) {
        // Pragmatic GFM-ish rule: treat '_' emphasis markers as disabled when they appear inside "words".
        // This avoids accidentally italicizing identifiers like foo_bar_baz.
        if (string.IsNullOrEmpty(text)) return false;
        int left = start - 1;
        int right = start + markerLength;
        if (left < 0 || right >= text.Length) return false;
        return char.IsLetterOrDigit(text[left]) && char.IsLetterOrDigit(text[right]);
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
        string inner = text.Substring(parenOpen + 1, parenClose - (parenOpen + 1));
        if (!TrySplitUrlAndOptionalTitle(inner, out href, out title)) {
            href = inner.Trim();
            title = null;
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
        string inner = text.Substring(altEnd + 2, imgClose - (altEnd + 2));
        if (!TrySplitUrlAndOptionalTitle(inner, out img, out imgTitle)) {
            img = inner.Trim();
            imgTitle = null;
        }
        int closeBracket = (imgClose + 1 < text.Length) ? text.IndexOf(']', imgClose + 1) : -1;
        if (closeBracket != imgClose + 1) return false;
        int parenOpen2 = (closeBracket + 1 < text.Length && text[closeBracket + 1] == '(') ? closeBracket + 1 : -1;
        if (parenOpen2 != closeBracket + 1) return false;
        int parenClose2 = FindMatchingParen(text, parenOpen2);
        if (parenClose2 < 0) return false;
        href = text.Substring(parenOpen2 + 1, parenClose2 - (parenOpen2 + 1)).Trim();
        consumed = parenClose2 - start + 1;
        return true;
    }

    private static bool TrySplitUrlAndOptionalTitle(string? inner, out string url, out string? title) {
        url = string.Empty;
        title = null;
        if (inner == null) return false;
        if (string.IsNullOrWhiteSpace(inner)) return false;

        var t = inner.Trim();
        if (t.Length == 0) return false;

        // CommonMark: destination can be wrapped in <...> to allow spaces and parentheses safely.
        if (t[0] == '<') {
            int gt = t.IndexOf('>');
            if (gt > 1) {
                url = t.Substring(1, gt - 1).Trim();
                var rest = t.Substring(gt + 1).Trim();
                if (rest.Length > 0) title = TryParseOptionalTitleToken(rest);
                return true;
            }
        }

        int ws = IndexOfWhitespace(t);
        if (ws < 0) { url = t; title = null; return true; }

        url = t.Substring(0, ws).Trim();
        var remaining = t.Substring(ws).Trim();
        if (remaining.Length == 0) { title = null; return true; }

        title = TryParseOptionalTitleToken(remaining);
        return true;
    }

    private static int IndexOfWhitespace(string s) {
        for (int i = 0; i < s.Length; i++) if (char.IsWhiteSpace(s[i])) return i;
        return -1;
    }

    private static string? TryParseOptionalTitleToken(string s) {
        if (string.IsNullOrWhiteSpace(s)) return null;
        var t = s.Trim();
        if (t.Length < 2) return null;
        if ((t[0] == '"' && t[t.Length - 1] == '"') ||
            (t[0] == '\'' && t[t.Length - 1] == '\'') ||
            (t[0] == '(' && t[t.Length - 1] == ')')) {
            return t.Substring(1, t.Length - 2);
        }
        return null;
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
        string inner = text.Substring(altEnd + 2, parenClose - (altEnd + 2));
        if (!TrySplitUrlAndOptionalTitle(inner, out src, out title)) {
            src = inner.Trim();
            title = null;
        }
        consumed = parenClose - start + 1;
        return true;
    }

    private static bool TryParseReferenceImage(string text, int start, out int consumed, out string alt, out string label) {
        consumed = 0; alt = label = string.Empty;
        if (start + 1 >= text.Length || text[start] != '!' || text[start + 1] != '[') return false;
        int altEnd = text.IndexOf(']', start + 2);
        if (altEnd < 0) return false;

        alt = text.Substring(start + 2, altEnd - (start + 2));

        // Inline image uses "(...)" and is handled elsewhere.
        if (altEnd + 1 < text.Length && text[altEnd + 1] == '(') return false;

        // Full or collapsed reference: ![alt][label] or ![alt][]
        if (altEnd + 1 < text.Length && text[altEnd + 1] == '[') {
            int labelEnd = text.IndexOf(']', altEnd + 2);
            if (labelEnd < 0) return false;
            label = text.Substring(altEnd + 2, labelEnd - (altEnd + 2));
            if (string.IsNullOrEmpty(label)) label = alt;
            consumed = labelEnd - start + 1;
            return true;
        }

        // Shortcut: ![label]
        label = alt;
        consumed = altEnd - start + 1;
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
        // Require a boundary on the left so we don't linkify inside longer words.
        if (start > 0 && char.IsLetterOrDigit(text[start - 1])) return false;
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

    private static bool StartsWithWww(string text, int start, out int end) {
        end = start;
        if (start + 4 > text.Length) return false;
        if (start > 0 && char.IsLetterOrDigit(text[start - 1])) return false;
        if (!(text.Substring(start).StartsWith("www.", StringComparison.OrdinalIgnoreCase))) return false;

        int i = start;
        while (i < text.Length) {
            char c = text[i];
            if (char.IsWhiteSpace(c)) break;
            if (c == ')' || c == ']' || c == '<') break;
            i++;
        }
        int scanEnd = i;
        while (i > start && (text[i - 1] == '.' || text[i - 1] == ',' || text[i - 1] == ';' || text[i - 1] == ':')) i--;

        // Must include at least one dot after the www.
        var token = text.Substring(start, i - start);
        if (token.Length <= 4) return false;
        if (token.IndexOf('.', 4) < 0) return false;

        // Right boundary: avoid linking as part of an identifier-like token.
        if (scanEnd < text.Length && IsEmailChar(text[scanEnd])) return false;

        end = i;
        return end > start + 4;
    }

    private static bool TryConsumePlainEmail(string text, int start, out int end, out string email) {
        end = start;
        email = string.Empty;
        if (start < 0 || start >= text.Length) return false;
        if (!IsEmailStartChar(text[start])) return false;
        if (start > 0 && IsEmailChar(text[start - 1])) return false;

        int i = start;
        bool sawAt = false;
        // Stop at whitespace or common "outside token" delimiters; keep it pragmatic.
        while (i < text.Length) {
            char c = text[i];
            if (char.IsWhiteSpace(c)) break;
            if (c == ')' || c == ']' || c == '<') break;
            if (!IsEmailChar(c)) break;
            if (c == '@') sawAt = true;
            i++;
        }
        if (!sawAt) return false;

        int scanEnd = i;
        int j = i;
        while (j > start && (text[j - 1] == '.' || text[j - 1] == ',' || text[j - 1] == ';' || text[j - 1] == ':')) j--;
        if (j <= start) return false;

        var token = text.Substring(start, j - start);
        if (!LooksLikeEmail(token)) return false;
        if (scanEnd < text.Length && IsEmailChar(text[scanEnd])) return false;

        end = j;
        email = token;
        return true;
    }

    private static bool IsEmailStartChar(char c) => char.IsLetterOrDigit(c);

    private static bool IsEmailChar(char c) {
        if (char.IsLetterOrDigit(c)) return true;
        return c == '@' || c == '.' || c == '-' || c == '_' || c == '+';
    }

    /// <summary>
    /// Parses a single line of Markdown inline content into a typed <see cref="InlineSequence"/>.
    /// This helper is exposed to allow other components (e.g., Word converter) to interpret
    /// inline markup in contexts like table cells where we currently store raw strings.
    /// </summary>
    /// <param name="text">Inline Markdown text.</param>
    /// <param name="options">Reader options controlling inline interpretation.</param>
    /// <returns>Parsed sequence of inline nodes.</returns>
    public static InlineSequence ParseInlineText(string? text, MarkdownReaderOptions? options = null) =>
        ParseInlineText(text, options, null);

    internal static InlineSequence ParseInlineText(string? text, MarkdownReaderOptions? options, MarkdownReaderState? state) =>
        ParseInlines(text ?? string.Empty, options ?? new MarkdownReaderOptions(), state);
}
